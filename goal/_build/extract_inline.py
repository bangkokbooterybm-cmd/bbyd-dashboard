"""Save a Google Drive download that came back INLINE (small files) to disk.

When download_file_content returns a small file, the JSON is shown inline instead of being
saved to a tool-results file. This script finds the newest such result for <fileId> in the
session transcripts and writes the decoded bytes.

usage: python3 -I extract_inline.py <fileId> <out.xlsx>
Prints the byte count, or exits 1 if nothing was found.
"""
import base64, glob, os, re, sys

file_id, out = sys.argv[1], sys.argv[2]
logs = sorted(glob.glob(os.path.expanduser('~/.claude/projects/*/*.jsonl')), key=os.path.getmtime, reverse=True)
pat = re.compile(r'\{\\*"content\\*":\\*"(UEsDB[A-Za-z0-9+/=]+)\\*",\\*"id\\*":\\*"' + re.escape(file_id))
for log in logs:  # newest transcript first; inside it, the last match is the newest download
    found = None
    with open(log, encoding='utf-8', errors='ignore') as f:
        for line in f:
            if file_id in line:
                for m in pat.finditer(line):
                    found = m.group(1)
    if found:
        data = base64.b64decode(found)
        if data[:2] != b'PK':
            sys.exit('not an xlsx (zip) payload')
        with open(out, 'wb') as f:
            f.write(data)
        print(len(data), 'bytes from', os.path.basename(log))
        sys.exit(0)
sys.exit('no inline download found for ' + file_id)
