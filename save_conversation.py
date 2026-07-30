import json
import sys
from datetime import datetime
from pathlib import Path

archive_path = Path(__file__).with_name('conversation_archive.jsonl')


def append_entry(role: str, text: str) -> None:
    entry = {
        'timestamp': datetime.now().isoformat(timespec='seconds'),
        'role': role,
        'text': text,
    }
    with archive_path.open('a', encoding='utf-8') as handle:
        handle.write(json.dumps(entry, ensure_ascii=False) + '\n')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python save_conversation.py <user|assistant> <text>')
        sys.exit(1)

    role = sys.argv[1]
    text = ' '.join(sys.argv[2:])
    append_entry(role, text)
    print(f'Saved to {archive_path}')
