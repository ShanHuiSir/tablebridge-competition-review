"""Export the synthetic acceptance case without any model/network call."""
import json
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from universal import DEMO_CODE, DEMO_SCHEMA
from universal_api import handle


def main(output):
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    session = {}
    with tempfile.TemporaryDirectory() as tmp:
        data = Path(tmp)
        handle('sample', {}, session, data)
        preview = handle('preview', {'code': DEMO_CODE, 'schema': DEMO_SCHEMA, 'confirmed': True}, session, data)
        assert preview['row_count'] == 5 and preview['issue_count'] == 1
        issue = preview['issues'][0]
        ready = handle('resolve', {'decisions': [{'action': 'edit', 'row_id': issue['row_id'], 'column': 'Mobile', 'value': '13800138001'}]}, session, data)
        assert ready['agent_status'] == 'ready_to_export'
        for format in ('xlsx', 'xls', 'csv'):
            result = handle('export', {'format': format, 'encoding': 'gb18030'}, session, data)
            for path in (data / result['run_id']).iterdir():
                if path.name == 'result.zip':
                    continue
                dest = output / (f'audit_{format}.json' if path.name == 'audit.json' else path.name)
                shutil.copyfile(path, dest)
        print(json.dumps({'output': str(output.resolve()), 'rows': 5, 'formats': ['xlsx', 'xls', 'csv/GB18030'], 'model_called': False}, ensure_ascii=False))


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else ROOT / 'docs' / '通用示例输出')
