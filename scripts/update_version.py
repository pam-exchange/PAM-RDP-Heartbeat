#!/usr/bin/env python3
import os
import re
import sys

def main():
    no_inc = '--no-increment' in sys.argv
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    version_file = os.path.join(root_dir, 'version')
    ahk_file = os.path.join(root_dir, 'src', 'pam-rdp-heartbeat.ahk')

    if os.path.exists(version_file):
        with open(version_file, 'r', encoding='utf-8') as f:
            v_str = f.read().strip()
    else:
        v_str = '1.1.0.0'

    if not no_inc:
        parts = v_str.split('.')
        if len(parts) == 4:
            parts[3] = str(int(parts[3]) + 1)
        elif len(parts) == 3:
            parts.append('1')
        else:
            parts = ['1', '1', '0', '1']
        new_version = '.'.join(parts)
        with open(version_file, 'w', encoding='utf-8') as f:
            f.write(new_version + '\n')
    else:
        new_version = v_str

    copyright_str = "Copyright (c) 2024-2026 PAM-Exchange"

    if os.path.exists(ahk_file):
        with open(ahk_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # Update gVersion:= "..."
        content = re.sub(
            r'gVersion\s*:=\s*"[^"]*"',
            f'gVersion:= "{new_version}"',
            content
        )

        # Update Copyright notice
        content = re.sub(
            r'Copyright\s*(?:©|\(c\))\s*\d{4}-\d{4}\s*PAM-Exchange',
            copyright_str,
            content
        )

        # Compiler directives
        directives = (
            f';@Ahk2Exe-SetFileVersion {new_version}\n'
            f';@Ahk2Exe-SetProductVersion {new_version}\n'
            f';@Ahk2Exe-SetCopyright {copyright_str}\n'
            f';@Ahk2Exe-SetCompanyName PAM-Exchange\n'
            f';@Ahk2Exe-SetProductName PAM RDP Heartbeat\n'
            f';@Ahk2Exe-SetDescription Sends heartbeat signal to RDP sessions\n'
        )

        # Remove existing directives if present
        content = re.sub(r'(?:\n*;@Ahk2Exe-[^\n]*)+', '', content)

        # Insert directives after top license comment block if present
        if content.startswith('/*'):
            end_comment_idx = content.find('*/')
            if end_comment_idx != -1:
                insert_pos = end_comment_idx + 2
                head = content[:insert_pos]
                tail = content[insert_pos:].lstrip('\r\n')
                content = head + '\n\n' + directives + '\n' + tail
            else:
                content = directives + '\n' + content.lstrip('\r\n')
        else:
            content = directives + '\n' + content.lstrip('\r\n')

        with open(ahk_file, 'w', encoding='utf-8') as f:
            f.write(content)

    print(new_version)

if __name__ == '__main__':
    main()
