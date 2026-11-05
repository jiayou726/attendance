"""Excel compound-document envelope for the legacy FDA/OLE DB importer.

Repair the BIFF8 WriteAccess record, and include generic Excel identification
streams that ordinary office applications write. xlwt's old username layout is
readable by xlrd but is rejected by the live importer. No office executable is
needed on the web server. The four metadata streams contain no business rows.
"""
import base64
from io import StringIO
import json
from pathlib import Path
import struct

from xlrd.compdoc import CompDoc

_METADATA = json.loads((Path(__file__).resolve().parents[1] /
                       "data/ftrace/xls_metadata.json").read_text())
_FREE, _END, _FAT = -1, -2, -3


def _normalise_write_access(workbook):
    # xlwt 1.3 emits the old fixed-width username layout, but BIFF8 requires
    # XLUnicodeString (MS-XLS 2.4.349). Two spaces mean no user name.
    result = bytearray(workbook)
    offset = 0
    while offset + 4 <= len(result):
        opcode, size = struct.unpack_from("<HH", result, offset)
        if opcode == 0x005C:
            if size != 112:
                raise ValueError("XLS WriteAccess 格式錯誤。")
            result[offset + 4:offset + 4 + size] = b"\x02\0\0  ".ljust(112, b" ")
            return bytes(result)
        if opcode == 0x000A:
            break
        offset += 4 + size
    raise ValueError("XLS 缺少 WriteAccess 資訊。")


def excel_envelope(xls_bytes):
    workbook = _normalise_write_access(
        CompDoc(xls_bytes, logfile=StringIO()).get_named_stream("Workbook"))
    # All BIFF workbooks from xlwt already occupy at least 4096 bytes.
    # Keeping that lower bound stores this stream in regular CFB sectors.
    workbook = workbook.ljust(max(4096, len(workbook)), b"\0")
    mini = bytearray()
    mini_fat = []
    streams = []
    for stream in _METADATA["streams"]:
        content = base64.b64decode(stream["base64"])
        start = len(mini) // 64
        count = (len(content) + 63) // 64
        mini.extend(content.ljust(count * 64, b"\0"))
        mini_fat.extend(start + i + 1 if i + 1 < count else _END for i in range(count))
        streams.append((stream["name"], start, len(content)))

    workbook_count = (len(workbook) + 511) // 512
    mini_count = (len(mini) + 511) // 512
    mini_sid = workbook_count
    mini_fat_sid = mini_sid + mini_count
    directory_sid = mini_fat_sid + 1
    fat_sid = directory_sid + 2
    fat_count = 1
    while (fat_sid + fat_count + 127) // 128 != fat_count:
        fat_count = (fat_sid + fat_count + 127) // 128
    if fat_count > 109 or len(mini_fat) > 128:
        raise ValueError("匯出資料超過 XLS 格式上限。")
    fat = [_FREE] * (fat_count * 128)
    for start, count in ((0, workbook_count), (mini_sid, mini_count),
                         (mini_fat_sid, 1), (directory_sid, 2)):
        for i in range(count):
            fat[start + i] = start + i + 1 if i + 1 < count else _END
    for i in range(fat_count):
        fat[fat_sid + i] = _FAT

    entries = [("Root Entry", mini_sid, len(mini)),
               ("Workbook", 0, len(workbook)), *streams]
    directory = bytearray()
    root_clsid = bytes.fromhex(_METADATA["root_clsid"])
    for i, (name, start, size) in enumerate(entries):
        node = _METADATA["directory"][i]
        if node["name"] != name:
            raise ValueError("XLS 格式設定錯誤。")
        encoded = (name + "\0").encode("utf-16le")
        directory.extend(encoded.ljust(64, b"\0"))
        directory.extend(struct.pack("<HBBiii16sIQQIQ", len(encoded),
            node["type"], node["colour"], node["left"], node["right"], node["child"],
            root_clsid if i == 0 else b"\0" * 16, 0, 0, 0, start, size))
    directory = directory.ljust(1024, b"\0")

    header = bytearray(512)
    header[:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<HHHHH", header, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", header, 40, 0, fat_count, directory_sid, 0,
                     4096, mini_fat_sid, 1, 0xFFFFFFFE, 0)
    struct.pack_into("<109i", header, 76,
        *([fat_sid + i for i in range(fat_count)] + [_FREE] * (109 - fat_count)))
    return (bytes(header) + workbook.ljust(workbook_count * 512, b"\0") +
            bytes(mini).ljust(mini_count * 512, b"\0") +
            struct.pack("<128i", *(mini_fat + [_FREE] * (128 - len(mini_fat)))) +
            bytes(directory) + struct.pack("<" + str(len(fat)) + "i", *fat))
