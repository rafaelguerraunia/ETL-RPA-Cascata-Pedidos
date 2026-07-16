"""Minimal MS-CFB (Compound File Binary Format) writer, sufficient to build
a vbaProject.bin. Implements the mini-stream/MiniFAT (mandatory per spec for
any stream smaller than 4096 bytes -- real readers, including Excel and
olefile, hardcode this cutoff and ignore the header field) plus regular FAT
sectors for larger streams. Directory sibling chains are simple sorted
right-leaning chains (a valid, if unbalanced, binary search tree -- CFB
readers do not validate red/black balance).

Usage:
    root = Storage("Root Entry")
    root.add_stream("PROJECT", project_bytes)
    vba = root.add_storage("VBA")
    vba.add_stream("dir", dir_bytes)
    ...
    data = build_cfb(root)
"""

import struct

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 4096
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF


def _cfb_name_key(name: str):
    # MS-CFB 2.6.4: compare by length first, then uppercase code unit by code unit.
    return (len(name), name.upper())


class Entry:
    def __init__(self, name, is_storage, data=b""):
        self.name = name
        self.is_storage = is_storage
        self.data = data
        self.children = []  # list of Entry (only for storages)
        self.dir_index = None

    def add_stream(self, name, data):
        e = Entry(name, is_storage=False, data=data)
        self.children.append(e)
        return e

    def add_storage(self, name):
        e = Entry(name, is_storage=True)
        self.children.append(e)
        return e


def Storage(name):
    return Entry(name, is_storage=True)


def _flatten(entry, out_list):
    out_list.append(entry)
    for c in sorted(entry.children, key=lambda e: _cfb_name_key(e.name)):
        _flatten(c, out_list)


def _build_sibling_chain(children_sorted_indices):
    links = {}
    for i, idx in enumerate(children_sorted_indices):
        right = children_sorted_indices[i + 1] if i + 1 < len(children_sorted_indices) else NOSTREAM
        links[idx] = (NOSTREAM, right)
    root_of_subtree = children_sorted_indices[0] if children_sorted_indices else NOSTREAM
    return root_of_subtree, links


def _pack_dir_entry(name, obj_type, color, left, right, child, clsid,
                     state_bits, ctime, mtime, start_sector, stream_size):
    name_utf16 = name.encode("utf-16-le")
    if len(name_utf16) > 62:
        raise ValueError(f"name too long: {name}")
    name_field = name_utf16 + b"\x00\x00"
    name_field = name_field + b"\x00" * (64 - len(name_field))
    name_len = len(name_utf16) + 2 if name else 0
    rec = struct.pack(
        "<64sHBBIII16sIQQIQ",
        name_field, name_len, obj_type, color,
        left, right, child,
        clsid, state_bits, ctime, mtime,
        start_sector, stream_size,
    )
    assert len(rec) == 128
    return rec


def build_cfb(root: Entry) -> bytes:
    root.name = "Root Entry"

    entries = []
    _flatten(root, entries)
    for i, e in enumerate(entries):
        e.dir_index = i
    n_entries = len(entries)
    n_entries_padded = ((n_entries + 3) // 4) * 4

    # ---- classify streams ----
    big_streams = []    # entries with size >= cutoff -> regular FAT
    mini_streams = []   # entries with 0 < size < cutoff -> mini-FAT / ministream
    for e in entries:
        if e.is_storage:
            continue
        size = len(e.data)
        if size == 0:
            continue
        if size >= MINI_STREAM_CUTOFF:
            big_streams.append(e)
        else:
            mini_streams.append(e)

    # ---- build ministream blob + mini sector chain ----
    mini_blob = bytearray()
    mini_chain_start = {}  # dir_index -> first mini sector index
    mini_fat = []          # one uint32 per mini sector (chain pointers)
    for e in mini_streams:
        size = len(e.data)
        num_mini_sectors = (size + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
        start = len(mini_blob) // MINI_SECTOR_SIZE
        padded = e.data + b"\x00" * (num_mini_sectors * MINI_SECTOR_SIZE - size)
        mini_blob.extend(padded)
        for k in range(num_mini_sectors):
            mini_fat.append((start + k + 1) if k + 1 < num_mini_sectors else ENDOFCHAIN)
        mini_chain_start[e.dir_index] = start

    ministream_size = len(mini_blob)
    minifat_sector_count = (len(mini_fat) * 4 + SECTOR_SIZE - 1) // SECTOR_SIZE if mini_fat else 0

    # ---- allocate regular sectors: [big streams][ministream][minifat][directory][FAT] ----
    sector_cursor = 0
    stream_sectors_data = []
    big_chain_start = {}

    for e in big_streams:
        size = len(e.data)
        num_sectors = (size + SECTOR_SIZE - 1) // SECTOR_SIZE
        start = sector_cursor
        padded = e.data + b"\x00" * (num_sectors * SECTOR_SIZE - size)
        for k in range(num_sectors):
            stream_sectors_data.append(padded[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE])
        sector_cursor += num_sectors
        big_chain_start[e.dir_index] = start
    total_big_sectors = sector_cursor

    # ministream sectors (regular FAT allocation for the concatenated mini blob)
    ministream_start_sector = sector_cursor
    ministream_num_sectors = (ministream_size + SECTOR_SIZE - 1) // SECTOR_SIZE if ministream_size else 0
    if ministream_num_sectors:
        padded = bytes(mini_blob) + b"\x00" * (ministream_num_sectors * SECTOR_SIZE - ministream_size)
        for k in range(ministream_num_sectors):
            stream_sectors_data.append(padded[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE])
        sector_cursor += ministream_num_sectors

    # minifat sectors
    minifat_start_sector = sector_cursor
    if minifat_sector_count:
        minifat_bytes = b"".join(struct.pack("<I", v & 0xFFFFFFFF) for v in mini_fat)
        minifat_bytes += b"\xff" * (minifat_sector_count * SECTOR_SIZE - len(minifat_bytes))
        for k in range(minifat_sector_count):
            stream_sectors_data.append(minifat_bytes[k * SECTOR_SIZE:(k + 1) * SECTOR_SIZE])
        sector_cursor += minifat_sector_count

    total_stream_sectors = sector_cursor  # big + ministream + minifat

    # ---- directory sectors ----
    dir_sector_count = n_entries_padded // 4
    dir_start_sector = sector_cursor
    sector_cursor += dir_sector_count
    total_before_fat = sector_cursor

    # ---- solve FAT sector count iteratively ----
    fat_sector_count = 1
    while True:
        total_sectors = total_before_fat + fat_sector_count
        needed = (total_sectors + 127) // 128
        if needed == fat_sector_count:
            break
        fat_sector_count = needed
    if fat_sector_count > 109:
        raise ValueError("FAT too large for header DIFAT (need DIFAT sectors, unimplemented)")

    fat_start_sector = total_before_fat
    total_sectors = total_before_fat + fat_sector_count

    # ---- build main FAT array ----
    fat = [FREESECT] * (fat_sector_count * 128)

    for e in big_streams:
        start = big_chain_start[e.dir_index]
        size = len(e.data)
        num_sectors = (size + SECTOR_SIZE - 1) // SECTOR_SIZE
        for k in range(num_sectors):
            sec = start + k
            fat[sec] = (start + k + 1) if k + 1 < num_sectors else ENDOFCHAIN

    for k in range(ministream_num_sectors):
        sec = ministream_start_sector + k
        fat[sec] = (ministream_start_sector + k + 1) if k + 1 < ministream_num_sectors else ENDOFCHAIN

    for k in range(minifat_sector_count):
        sec = minifat_start_sector + k
        fat[sec] = (minifat_start_sector + k + 1) if k + 1 < minifat_sector_count else ENDOFCHAIN

    for k in range(dir_sector_count):
        sec = dir_start_sector + k
        fat[sec] = (dir_start_sector + k + 1) if k + 1 < dir_sector_count else ENDOFCHAIN

    for k in range(fat_sector_count):
        fat[fat_start_sector + k] = FATSECT

    # ---- build directory entries tree (sibling links) ----
    left_right = {}
    child_of = {}

    def process(entry):
        if not entry.is_storage:
            return
        sorted_children = sorted(entry.children, key=lambda e: _cfb_name_key(e.name))
        idxs = [c.dir_index for c in sorted_children]
        root_idx, links = _build_sibling_chain(idxs)
        left_right.update(links)
        child_of[entry.dir_index] = root_idx
        for c in entry.children:
            process(c)

    process(root)

    # ---- serialize directory entries ----
    dir_bytes = bytearray()
    for e in entries:
        if e.dir_index == 0:
            obj_type = 5
        elif e.is_storage:
            obj_type = 1
        else:
            obj_type = 2
        color = 1

        left, right = left_right.get(e.dir_index, (NOSTREAM, NOSTREAM))
        child = child_of.get(e.dir_index, NOSTREAM) if e.is_storage else NOSTREAM

        if e.dir_index == 0:
            # Root Entry: StartSector/Size describe the ministream itself.
            start_sector = ministream_start_sector if ministream_num_sectors else ENDOFCHAIN
            stream_size = ministream_size
        elif e.is_storage:
            start_sector = 0
            stream_size = 0
        else:
            size = len(e.data)
            if size == 0:
                start_sector = ENDOFCHAIN
                stream_size = 0
            elif size >= MINI_STREAM_CUTOFF:
                start_sector = big_chain_start[e.dir_index]
                stream_size = size
            else:
                start_sector = mini_chain_start[e.dir_index]
                stream_size = size

        dir_bytes.extend(_pack_dir_entry(
            e.name, obj_type, color, left, right, child,
            b"\x00" * 16, 0, 0, 0, start_sector, stream_size,
        ))

    unused = _pack_dir_entry("", 0, 1, NOSTREAM, NOSTREAM, NOSTREAM, b"\x00" * 16, 0, 0, 0, 0, 0)
    while len(dir_bytes) < n_entries_padded * 128:
        dir_bytes.extend(unused)

    # ---- assemble file ----
    out = bytearray()

    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    header[8:24] = b"\x00" * 16
    struct.pack_into("<H", header, 24, 0x003E)
    struct.pack_into("<H", header, 26, 0x0003)
    struct.pack_into("<H", header, 28, 0xFFFE)
    struct.pack_into("<H", header, 30, 9)
    struct.pack_into("<H", header, 32, 6)
    header[34:40] = b"\x00" * 6
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, fat_sector_count)
    struct.pack_into("<I", header, 48, dir_start_sector)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, 0x00001000)  # mandatory 4096, per spec
    struct.pack_into("<I", header, 60, minifat_start_sector if minifat_sector_count else ENDOFCHAIN)
    struct.pack_into("<I", header, 64, minifat_sector_count)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)
    struct.pack_into("<I", header, 72, 0)
    difat = [FREESECT] * 109
    for k in range(fat_sector_count):
        difat[k] = fat_start_sector + k
    for k in range(109):
        struct.pack_into("<I", header, 76 + k * 4, difat[k])
    out.extend(header)

    for sec in stream_sectors_data:
        out.extend(sec)

    out.extend(dir_bytes)

    for k in range(fat_sector_count):
        block = fat[k * 128:(k + 1) * 128]
        for v in block:
            out.extend(struct.pack("<I", v & 0xFFFFFFFF))

    assert len(out) == 512 + total_sectors * SECTOR_SIZE, (len(out), total_sectors)
    return bytes(out)
