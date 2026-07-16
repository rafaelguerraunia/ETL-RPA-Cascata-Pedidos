"""MS-OVBA 2.4.1 'Compressed Container' codec (used for the 'dir' stream and
module source-code storage inside a VBA project). Implements a real (greedy)
LZ77-style compressor plus a decompressor, so we can round-trip test locally
without any external MS Office / oletools dependency for correctness.
"""

def _ceil_log2(x: int) -> int:
    if x <= 1:
        return 0
    return (x - 1).bit_length()


def _bit_count(difference: int) -> int:
    return max(_ceil_log2(difference), 4)


def _find_match(data: bytes, pos: int, chunk_start: int, chunk_end: int):
    """Greedy longest-match search within the current chunk's decompressed
    window [chunk_start, pos). Allows overlapping matches (offset < length).
    Returns (length, offset) or None."""
    difference = pos - chunk_start
    if difference <= 0:
        return None
    bc = _bit_count(difference)
    max_offset = 1 << bc
    max_offset = min(max_offset, difference)
    length_bits = 16 - bc
    max_length = (1 << length_bits) - 1 + 3
    max_length = min(max_length, chunk_end - pos)
    if max_length < 3:
        return None

    best_len = 0
    best_off = 0
    # search all candidate offsets (small data sizes here, O(n*window) is fine)
    lo = max(chunk_start, pos - max_offset)
    for cand in range(pos - 1, lo - 1, -1):
        offset = pos - cand
        if offset > max_offset:
            continue
        length = 0
        limit = max_length
        while length < limit and data[cand + (length % offset)] == data[pos + length]:
            length += 1
        if length > best_len:
            best_len = length
            best_off = offset
            if best_len >= max_length:
                break
    if best_len >= 3:
        return best_len, best_off
    return None


def _compress_chunk_tokens(data: bytes, chunk_start: int, chunk_end: int) -> bytes:
    """Encode data[chunk_start:chunk_end] using the token-sequence format
    (FlagByte + up to 8 literal/copy tokens per group)."""
    out = bytearray()
    pos = chunk_start
    group_flags = 0
    group_tokens = bytearray()
    group_count = 0

    def flush_group():
        nonlocal group_flags, group_tokens, group_count
        if group_count > 0:
            out.append(group_flags)
            out.extend(group_tokens)
        group_flags = 0
        group_tokens = bytearray()
        group_count = 0

    while pos < chunk_end:
        match = _find_match(data, pos, chunk_start, chunk_end)
        if match is not None:
            length, offset = match
            difference = pos - chunk_start
            bc = _bit_count(difference)
            token = ((offset - 1) << (16 - bc)) | ((length - 3) & ((1 << (16 - bc)) - 1))
            group_tokens.append(token & 0xFF)
            group_tokens.append((token >> 8) & 0xFF)
            group_flags |= (1 << group_count)
            pos += length
        else:
            group_tokens.append(data[pos])
            pos += 1
        group_count += 1
        if group_count == 8:
            flush_group()
    flush_group()
    return bytes(out)


def compress(raw: bytes) -> bytes:
    out = bytearray()
    out.append(0x01)  # SignatureByte
    n = len(raw)
    pos = 0
    if n == 0:
        return bytes(out)
    while pos < n:
        chunk_start = pos
        chunk_end = min(pos + 4096, n)
        chunk_len = chunk_end - chunk_start

        if chunk_len == 4096:
            # try token-compressed form first; only valid if it fits in 4096 bytes
            token_data = _compress_chunk_tokens(raw, chunk_start, chunk_end)
            if len(token_data) < 4096:
                header = (len(token_data) + 2 - 3) & 0x0FFF
                header |= (0b011 << 12)
                header |= (1 << 15)
                out.append(header & 0xFF)
                out.append((header >> 8) & 0xFF)
                out.extend(token_data)
            else:
                # raw (uncompressed) chunk: exactly 4096 data bytes
                header = (4096 + 2 - 3) & 0x0FFF
                header |= (0b011 << 12)
                header |= (0 << 15)
                out.append(header & 0xFF)
                out.append((header >> 8) & 0xFF)
                out.extend(raw[chunk_start:chunk_end])
        else:
            # final, partial chunk: MUST use token-compressed form
            token_data = _compress_chunk_tokens(raw, chunk_start, chunk_end)
            assert len(token_data) <= 4096, "final chunk overflowed 4096-byte compressed budget"
            header = (len(token_data) + 2 - 3) & 0x0FFF
            header |= (0b011 << 12)
            header |= (1 << 15)
            out.append(header & 0xFF)
            out.append((header >> 8) & 0xFF)
            out.extend(token_data)

        pos = chunk_end
    return bytes(out)


def decompress(compressed: bytes) -> bytes:
    if len(compressed) == 0:
        return b""
    assert compressed[0] == 0x01, "bad CompressedContainer signature byte"
    out = bytearray()
    i = 1
    n = len(compressed)
    while i < n:
        header = compressed[i] | (compressed[i + 1] << 8)
        i += 2
        chunk_size = (header & 0x0FFF) + 3  # total bytes incl. 2-byte header
        flag = (header >> 15) & 0x1
        data_len = chunk_size - 2
        chunk_data = compressed[i:i + data_len]
        i += data_len

        chunk_start = len(out)
        if flag == 0:
            assert len(chunk_data) == 4096, "raw chunk must be exactly 4096 bytes"
            out.extend(chunk_data)
        else:
            j = 0
            dl = len(chunk_data)
            while j < dl:
                flags = chunk_data[j]
                j += 1
                for bit in range(8):
                    if j >= dl:
                        break
                    if (flags >> bit) & 1:
                        token = chunk_data[j] | (chunk_data[j + 1] << 8)
                        j += 2
                        difference = len(out) - chunk_start
                        bc = _bit_count(difference)
                        length_mask = 0xFFFF >> bc
                        offset_mask = (~length_mask) & 0xFFFF
                        length = (token & length_mask) + 3
                        offset = ((token & offset_mask) >> (16 - bc)) + 1
                        start = len(out) - offset
                        for k in range(length):
                            out.append(out[start + k])
                    else:
                        out.append(chunk_data[j])
                        j += 1
    return bytes(out)


if __name__ == "__main__":
    import os, random

    tests = [
        b"",
        b"A",
        b"Attribute VB_Name = \"TestModule\"\r\nSub HelloWorld()\r\n    MsgBox \"Hi\"\r\nEnd Sub\r\n",
        b"x" * 4096,
        b"x" * 4097,
        b"x" * 8192,
        b"x" * 8193,
        (b"ABCDEFGH" * 600)[:4096],
        (b"ABCDEFGH" * 600)[:4095],
        (b"ABCDEFGH" * 600)[:3642],
        (b"ABCDEFGH" * 600)[:4090],
        os.urandom(100),
        os.urandom(2000),
    ]
    random.seed(42)
    big = bytearray()
    words = [b"Dim ", b"Set ", b"End Sub\r\n", b"    ", b"VBELN", b"clsFiltro", b"resultado", b"\r\n"]
    for _ in range(20000):
        big.extend(random.choice(words))
    tests.append(bytes(big))

    for idx, t in enumerate(tests):
        c = compress(t)
        d = decompress(c)
        ok = d == t
        print(f"test {idx}: len={len(t):6d} compressed={len(c):6d} roundtrip={'OK' if ok else 'FAIL'}")
        if not ok:
            raise SystemExit(f"ROUND TRIP FAILED on test {idx}")

    print("ALL ROUND TRIP TESTS PASSED")
