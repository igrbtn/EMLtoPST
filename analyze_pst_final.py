#!/usr/bin/env python3
"""
Analyze PST: parse NBT, read message classes.

This PST uses a non-standard PC format:
- bClientSig = 0xBC (instead of standard 0x6C)
- No BTH header (0xB5) wrapper
- Properties stored as flat 8-byte entries directly in HN allocations
- Each entry: propId(2) + propType(2) + dwValue(4)
- String values are HNIDs referencing other allocations
- HID index is offset by +1 because allocation [0] contains the HNHDR itself
  So HID index i -> rgibAlloc[i]..rgibAlloc[i+1] (not rgibAlloc[i-1]..rgibAlloc[i])
"""
import struct
import sys
from collections import Counter, defaultdict
import datetime

PST_PATH = "/Users/igorbatin/Documents/VaibPro/EMLtoPST/MockOrg Taylor Karen.pst"

PROP_NAMES = {
    0x0017: "PR_IMPORTANCE",
    0x001A: "PR_MESSAGE_CLASS",
    0x0026: "PR_PRIORITY",
    0x0036: "PR_SENSITIVITY",
    0x0037: "PR_SUBJECT",
    0x0039: "PR_CLIENT_SUBMIT_TIME",
    0x003D: "PR_SUBJECT_PREFIX",
    0x0042: "PR_SENT_REPRESENTING_NAME",
    0x0064: "PR_SENT_REPRESENTING_ADDRTYPE",
    0x0065: "PR_SENT_REPRESENTING_EMAIL",
    0x0070: "PR_CONVERSATION_TOPIC",
    0x0C1A: "PR_SENDER_NAME",
    0x0C1E: "PR_SENDER_ADDRTYPE",
    0x0C1F: "PR_SENDER_EMAIL_ADDRESS",
    0x0E04: "PR_DISPLAY_TO",
    0x0E06: "PR_MESSAGE_DELIVERY_TIME",
    0x0E07: "PR_MESSAGE_FLAGS",
    0x0E08: "PR_MESSAGE_SIZE",
    0x0E1B: "PR_HASATTACH",
    0x1000: "PR_BODY",
    0x3001: "PR_DISPLAY_NAME",
    0x3007: "PR_CREATION_TIME",
    0x3008: "PR_LAST_MODIFICATION_TIME",
    0x3613: "PR_CONTAINER_CLASS",
    0x3602: "PR_CONTENT_COUNT",
    0x3603: "PR_CONTENT_UNREAD_COUNT",
    0x360A: "PR_SUBFOLDERS",
}

NID_TYPE_NAMES = {
    0x00: "NID_TYPE_HID",
    0x01: "NID_TYPE_INTERNAL",
    0x02: "NID_TYPE_NORMAL_FOLDER",
    0x03: "NID_TYPE_SEARCH_FOLDER",
    0x04: "NID_TYPE_NORMAL_MESSAGE",
    0x05: "NID_TYPE_ATTACHMENT",
    0x06: "NID_TYPE_SEARCH_UPDATE_QUEUE",
    0x07: "NID_TYPE_SEARCH_CRITERIA_OBJECT",
    0x08: "NID_TYPE_ASSOC_MESSAGE",
    0x0A: "NID_TYPE_CONTENTS_TABLE_INDEX",
    0x0B: "NID_TYPE_RECEIVE_FOLDER_TABLE",
    0x0C: "NID_TYPE_OUTGOING_QUEUE_TABLE",
    0x0D: "NID_TYPE_HIERARCHY_TABLE",
    0x0E: "NID_TYPE_CONTENTS_TABLE",
    0x0F: "NID_TYPE_ASSOC_CONTENTS_TABLE",
    0x10: "NID_TYPE_SEARCH_CONTENTS_TABLE",
    0x11: "NID_TYPE_ATTACHMENT_TABLE",
    0x12: "NID_TYPE_RECIPIENT_TABLE",
    0x13: "NID_TYPE_SEARCH_TABLE_INDEX",
    0x1F: "NID_TYPE_LTP",
}


class PSTReader:
    def __init__(self, path):
        self.path = path
        self.f = open(path, "rb")
        self._read_header()
        self._build_bbt_index()

    def _read_header(self):
        self.f.seek(0)
        hdr = self.f.read(564)
        self.wVer = struct.unpack_from('<H', hdr, 10)[0]
        ROOT_OFF = 0xB4
        self.nbt_root_ib = struct.unpack_from('<Q', hdr, ROOT_OFF + 0x2C)[0]
        self.bbt_root_ib = struct.unpack_from('<Q', hdr, ROOT_OFF + 0x3C)[0]

        # Get tree levels
        page = self._read_page(self.nbt_root_ib)
        self.nbt_level = page[491]

        page = self._read_page(self.bbt_root_ib)
        self.bbt_level = page[491]

    def _read_page(self, ib):
        self.f.seek(ib)
        return self.f.read(512)

    def _read_block(self, ib, cb):
        self.f.seek(ib)
        return self.f.read(cb)

    def _build_bbt_index(self):
        """Build complete BBT index: bid -> {ib, cb}."""
        self.bbt = {}
        self._walk_bbt(self.bbt_root_ib)

    def _walk_bbt(self, ib):
        page = self._read_page(ib)
        cEnt = page[488]
        cbEnt = page[490]
        cLevel = page[491]

        if cLevel == 0:
            for i in range(cEnt):
                off = i * cbEnt
                bid = struct.unpack_from('<Q', page, off)[0]
                entry_ib = struct.unpack_from('<Q', page, off + 8)[0]
                entry_cb = struct.unpack_from('<H', page, off + 16)[0]
                self.bbt[bid] = (entry_ib, entry_cb)
        else:
            for i in range(cEnt):
                off = i * 24
                child_ib = struct.unpack_from('<Q', page, off + 16)[0]
                self._walk_bbt(child_ib)

    def walk_nbt(self):
        """Yield all NBT entries."""
        yield from self._walk_nbt(self.nbt_root_ib)

    def _walk_nbt(self, ib):
        page = self._read_page(ib)
        cEnt = page[488]
        cbEnt = page[490]
        cLevel = page[491]

        if cLevel == 0:
            for i in range(cEnt):
                off = i * cbEnt
                nid = struct.unpack_from('<Q', page, off)[0] & 0xFFFFFFFF
                bidData = struct.unpack_from('<Q', page, off + 8)[0]
                bidSub = struct.unpack_from('<Q', page, off + 16)[0]
                nidParent = struct.unpack_from('<I', page, off + 24)[0]
                yield {
                    'nid': nid,
                    'nid_type': nid & 0x1F,
                    'bidData': bidData,
                    'bidSub': bidSub,
                    'nidParent': nidParent,
                }
        else:
            for i in range(cEnt):
                off = i * 24
                child_ib = struct.unpack_from('<Q', page, off + 16)[0]
                yield from self._walk_nbt(child_ib)

    def read_pc_properties(self, bid):
        """Read properties from a PC stored at the given BID.
        Returns dict: propId -> (propType, resolved_value)
        """
        if bid == 0 or bid not in self.bbt:
            return None

        block_ib, block_cb = self.bbt[bid]
        data = self._read_block(block_ib, block_cb)

        if len(data) < 12:
            return None

        # Parse HN header
        ibHnpm = struct.unpack_from('<H', data, 0)[0]
        bSig = data[2]

        if bSig != 0xEC:
            return None

        hidUserRoot = struct.unpack_from('<I', data, 4)[0]

        if ibHnpm >= len(data) or ibHnpm + 2 > len(data):
            return None

        # Parse page map
        cAlloc = struct.unpack_from('<H', data, ibHnpm)[0]
        if ibHnpm + 2 + (cAlloc + 1) * 2 > len(data):
            return None

        allocs = []
        for i in range(cAlloc + 1):
            allocs.append(struct.unpack_from('<H', data, ibHnpm + 2 + i * 2)[0])

        # Get the allocation for hidUserRoot
        # hidUserRoot uses standard MS-PST mapping: hidIndex -> allocs[hidIndex-1]..allocs[hidIndex]
        hidIndex = (hidUserRoot >> 5) & 0x7FF
        if hidIndex == 0 or hidIndex > cAlloc:
            return None

        prop_start = allocs[hidIndex - 1]
        prop_end = allocs[hidIndex]

        if prop_start >= prop_end:
            return None

        # Parse property entries (8 bytes each)
        props = {}
        num_entries = (prop_end - prop_start) // 8
        for i in range(num_entries):
            ent_off = prop_start + i * 8
            if ent_off + 8 > len(data):
                break
            propId = struct.unpack_from('<H', data, ent_off)[0]
            propType = struct.unpack_from('<H', data, ent_off + 2)[0]
            dwValue = struct.unpack_from('<I', data, ent_off + 4)[0]

            # Resolve value based on type
            if propType in (0x001F, 0x001E):  # PT_UNICODE, PT_STRING8
                val = self._resolve_hnid_string(data, allocs, cAlloc, dwValue, propType)
                props[propId] = (propType, val)
            elif propType == 0x0003:  # PT_LONG
                props[propId] = (propType, dwValue)
            elif propType == 0x000B:  # PT_BOOLEAN
                props[propId] = (propType, bool(dwValue))
            elif propType == 0x0040:  # PT_SYSTIME
                val = self._resolve_hnid_time(data, allocs, cAlloc, dwValue)
                props[propId] = (propType, val)
            elif propType == 0x0102:  # PT_BINARY
                val = self._resolve_hnid_binary(data, allocs, cAlloc, dwValue)
                props[propId] = (propType, val)
            else:
                props[propId] = (propType, dwValue)

        return props

    def _resolve_hnid_string(self, data, allocs, cAlloc, hnid, propType):
        """Resolve an HNID to a string value.
        In this PST, HID index i maps to rgibAlloc[i]..rgibAlloc[i+1]
        (offset by 1 because alloc[0] = HNHDR).
        """
        if hnid == 0:
            return ""

        hidIndex = (hnid >> 5) & 0x7FF
        if hidIndex == 0 or hidIndex >= cAlloc:
            # Maybe inline value (4 bytes max)
            try:
                raw = hnid.to_bytes(4, 'little')
                if propType == 0x001F:
                    return raw.decode('utf-16-le', errors='replace').rstrip('\x00')
                else:
                    return raw.decode('ascii', errors='replace').rstrip('\x00')
            except:
                return f"<hnid:0x{hnid:08x}>"

        s = allocs[hidIndex]
        e = allocs[hidIndex + 1]
        if s >= e or e > len(data):
            return f"<empty:idx{hidIndex}>"

        raw = data[s:e]
        if propType == 0x001F:
            return raw.decode('utf-16-le', errors='replace').rstrip('\x00')
        else:
            return raw.decode('ascii', errors='replace').rstrip('\x00')

    def _resolve_hnid_time(self, data, allocs, cAlloc, hnid):
        """Resolve an HNID to a datetime value."""
        if hnid == 0:
            return None

        hidIndex = (hnid >> 5) & 0x7FF
        if hidIndex == 0 or hidIndex >= cAlloc:
            return None

        s = allocs[hidIndex]
        e = allocs[hidIndex + 1]
        if e - s != 8 or e > len(data):
            return None

        ft = struct.unpack_from('<Q', data, s)[0]
        EPOCH_DIFF = 116444736000000000
        if ft > EPOCH_DIFF:
            ts = (ft - EPOCH_DIFF) / 10000000
            try:
                return datetime.datetime.utcfromtimestamp(ts)
            except:
                return None
        return None

    def _resolve_hnid_binary(self, data, allocs, cAlloc, hnid):
        """Resolve an HNID to binary data."""
        if hnid == 0:
            return b''

        hidIndex = (hnid >> 5) & 0x7FF
        if hidIndex == 0 or hidIndex >= cAlloc:
            return hnid.to_bytes(4, 'little')

        s = allocs[hidIndex]
        e = allocs[hidIndex + 1]
        if s >= e or e > len(data):
            return b''

        return data[s:e]

    def close(self):
        self.f.close()


def main():
    pst = PSTReader(PST_PATH)

    print(f"PST: {PST_PATH}")
    print(f"NBT root: 0x{pst.nbt_root_ib:x} (level {pst.nbt_level})")
    print(f"BBT root: 0x{pst.bbt_root_ib:x} (level {pst.bbt_level})")
    print(f"BBT entries: {len(pst.bbt)}")

    # Collect all NBT entries
    all_entries = list(pst.walk_nbt())
    print(f"\nTotal NBT entries: {len(all_entries)}")

    # Categorize by type
    type_counts = Counter()
    type_entries = defaultdict(list)
    for entry in all_entries:
        t = entry['nid_type']
        type_counts[t] += 1
        type_entries[t].append(entry)

    print(f"\n{'='*60}")
    print(f"NID TYPE DISTRIBUTION")
    print(f"{'='*60}")
    for t in sorted(type_counts.keys()):
        name = NID_TYPE_NAMES.get(t, f"UNKNOWN_0x{t:02x}")
        print(f"  0x{t:02x} {name:40s}: {type_counts[t]}")

    # Message NIDs
    msg_entries = type_entries[0x04]
    print(f"\n{'='*60}")
    print(f"MESSAGE NIDS (type 0x04): {len(msg_entries)}")
    print(f"{'='*60}")

    print(f"\nFirst 5:")
    for e in msg_entries[:5]:
        print(f"  NID=0x{e['nid']:08x}, bidData=0x{e['bidData']:x}, parent=0x{e['nidParent']:08x}")
    print(f"Last 5:")
    for e in msg_entries[-5:]:
        print(f"  NID=0x{e['nid']:08x}, bidData=0x{e['bidData']:x}, parent=0x{e['nidParent']:08x}")

    # Map parent folder NIDs to folder info
    folder_entries = type_entries[0x02]
    folder_names = {}
    for fe in folder_entries:
        props = pst.read_pc_properties(fe['bidData'])
        if props and 0x3001 in props:
            folder_names[fe['nid']] = props[0x3001][1]
        else:
            folder_names[fe['nid']] = f"<folder 0x{fe['nid']:08x}>"

    print(f"\n{'='*60}")
    print(f"FOLDERS")
    print(f"{'='*60}")
    for fe in folder_entries:
        name = folder_names.get(fe['nid'], '???')
        # Count messages in this folder
        msg_count = sum(1 for m in msg_entries if m['nidParent'] == fe['nid'])
        print(f"  NID=0x{fe['nid']:08x}, name='{name}', messages={msg_count}, parent=0x{fe['nidParent']:08x}")

    # Read PR_MESSAGE_CLASS for all messages
    print(f"\n{'='*60}")
    print(f"READING PR_MESSAGE_CLASS FOR ALL {len(msg_entries)} MESSAGES")
    print(f"{'='*60}")

    msg_class_counter = Counter()
    msg_class_by_folder = defaultdict(Counter)
    msg_class_examples = defaultdict(list)
    errors = Counter()
    all_prop_ids_seen = Counter()

    for idx, entry in enumerate(msg_entries):
        if idx % 250 == 0 and idx > 0:
            print(f"  Progress: {idx}/{len(msg_entries)}...", file=sys.stderr)

        props = pst.read_pc_properties(entry['bidData'])
        if props is None:
            errors['no_props'] += 1
            continue

        # Track all property IDs seen
        for pid in props:
            all_prop_ids_seen[pid] += 1

        # Get message class
        if 0x001A in props:
            msg_class = props[0x001A][1]
            msg_class_counter[msg_class] += 1

            folder_name = folder_names.get(entry['nidParent'], f"0x{entry['nidParent']:08x}")
            msg_class_by_folder[folder_name][msg_class] += 1

            if len(msg_class_examples[msg_class]) < 3:
                msg_class_examples[msg_class].append(entry)
        else:
            errors['no_message_class'] += 1

    print(f"\n{'='*60}")
    print(f"MESSAGE CLASS DISTRIBUTION")
    print(f"{'='*60}")
    total_resolved = sum(msg_class_counter.values())
    for mc, count in msg_class_counter.most_common():
        pct = count / len(msg_entries) * 100
        print(f"  {mc}: {count} ({pct:.1f}%)")
        examples = msg_class_examples[mc]
        for ex in examples[:2]:
            folder = folder_names.get(ex['nidParent'], '?')
            # Get subject
            props = pst.read_pc_properties(ex['bidData'])
            subj = props.get(0x0037, (0, '<no subject>'))[1] if props else '?'
            if isinstance(subj, str) and len(subj) > 60:
                subj = subj[:57] + "..."
            print(f"    Example: NID=0x{ex['nid']:08x}, folder='{folder}', subject='{subj}'")

    print(f"\n{'='*60}")
    print(f"MESSAGE CLASSES BY FOLDER")
    print(f"{'='*60}")
    for folder_name in sorted(msg_class_by_folder.keys()):
        classes = msg_class_by_folder[folder_name]
        total_in_folder = sum(classes.values())
        print(f"\n  Folder: '{folder_name}' ({total_in_folder} messages)")
        for mc, cnt in classes.most_common():
            print(f"    {mc}: {cnt}")

    print(f"\n{'='*60}")
    print(f"PROPERTY IDS SEEN ACROSS ALL MESSAGES")
    print(f"{'='*60}")
    for pid, cnt in all_prop_ids_seen.most_common():
        name = PROP_NAMES.get(pid, f"0x{pid:04x}")
        print(f"  0x{pid:04x} ({name}): {cnt}/{total_resolved} messages")

    print(f"\n{'='*60}")
    print(f"ERRORS / UNRESOLVED")
    print(f"{'='*60}")
    for err, cnt in errors.most_common():
        print(f"  {err}: {cnt}")
    print(f"  Total resolved: {total_resolved}/{len(msg_entries)}")
    print(f"  Total errors: {sum(errors.values())}/{len(msg_entries)}")

    # Step: Examine sample of each message class in detail
    print(f"\n{'='*60}")
    print(f"DETAILED PROPERTIES FOR ONE MESSAGE OF EACH CLASS")
    print(f"{'='*60}")
    for mc in msg_class_counter:
        examples = msg_class_examples[mc]
        if examples:
            ex = examples[0]
            props = pst.read_pc_properties(ex['bidData'])
            if props:
                print(f"\n  === {mc} (NID=0x{ex['nid']:08x}) ===")
                for pid in sorted(props.keys()):
                    ptype, pval = props[pid]
                    name = PROP_NAMES.get(pid, f"0x{pid:04x}")
                    val_str = str(pval)
                    if isinstance(pval, str) and len(pval) > 80:
                        val_str = pval[:77] + "..."
                    elif isinstance(pval, bytes) and len(pval) > 32:
                        val_str = pval[:32].hex() + "..."
                    print(f"    0x{pid:04x} ({name}): type=0x{ptype:04x}, val={val_str}")

    # Check for unusual NID patterns
    print(f"\n{'='*60}")
    print(f"UNUSUAL NID PATTERNS")
    print(f"{'='*60}")

    # Check NID spacing
    nids = sorted(e['nid'] for e in msg_entries)
    diffs = [nids[i+1] - nids[i] for i in range(len(nids)-1)]
    diff_counter = Counter(diffs)
    print(f"\n  NID spacing distribution:")
    for diff, cnt in diff_counter.most_common(5):
        print(f"    Spacing 0x{diff:x} ({diff}): {cnt} times")

    print(f"\n  NID range: 0x{nids[0]:08x} to 0x{nids[-1]:08x}")

    # Check for gaps in parent folder assignment
    parent_counter = Counter(e['nidParent'] for e in msg_entries)
    print(f"\n  Parent folder distribution:")
    for parent, cnt in parent_counter.most_common():
        name = folder_names.get(parent, f"0x{parent:08x}")
        print(f"    {name} (0x{parent:08x}): {cnt} messages")

    pst.close()
    print(f"\nDone.")


if __name__ == "__main__":
    main()
