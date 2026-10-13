"""State transitions for the general agent; exported runs are immutable snapshots."""
import base64
import copy
import hashlib
import io
import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from core import ContractError
from universal import (
    ROOT, DEMO_CODE, DEMO_PROMPT, DEMO_SCHEMA, read_source, default_schema,
    validate_schema, model_plan, run_sandbox, issues_for, apply_decisions,
    public_result, export_rows, verify_export,
)


def state(session):
    return session.setdefault('universal', {'tables': {}, 'raw': {}, 'plan': None, 'result': None})


def fingerprint(tables):
    return {k: t['headers'] for k, t in tables.items()}


def public_state(s):
    return {
        'tables': {k: {'filename': t['filename'], 'sheet': t['sheet'], 'headers': t['headers'], 'row_count': len(t['rows']), 'sha256': t['sha256'], 'header_row': t['header_row'], 'preview': [{'row': r['row'], 'values': {h: str(v) if v is not None else '' for h, v in r['values'].items()}} for r in t['rows'] if not r['blank']][:6]} for k, t in s['tables'].items()},
        'schema': s.get('schema'), 'plan': s['plan'],
    }


def template_dir(data):
    path = data / 'templates'
    path.mkdir(parents=True, exist_ok=True)
    return path


def templates(data):
    result = []
    for path in sorted(template_dir(data).glob('*.json')):
        try:
            value = json.loads(path.read_text(encoding='utf-8'))
            result.append({k: value[k] for k in ('id', 'name', 'created_at', 'summary', 'source_run')})
        except (OSError, ValueError, KeyError):
            continue
    return result


def required(s):
    if not s['tables'].get('source1') or not s['tables'].get('target'):
        raise ContractError('至少上传源表 A 和目标模板')


def handle(path, body, session, data):
    s = state(session)
    if path == 'templates':
        return {'templates': templates(data)}
    if path == 'sample':
        names = {'source1': '系统A_员工报销.xlsx', 'source2': '系统B_员工费用.csv', 'target': '目标HR_员工导入.xlsx'}
        s['tables'] = {}
        s['raw'] = {}
        for role, filename in names.items():
            raw = (ROOT / 'samples' / 'universal' / filename).read_bytes()
            s['tables'][role] = read_source(raw, filename)
            s['raw'][role] = raw
        s['schema'] = copy.deepcopy(DEMO_SCHEMA)
        s['plan'] = {'summary': '合并两个系统，清洗电话与日期，剔除负金额、去重并拆分地址。异常等待人工决策。', 'code': DEMO_CODE, 'schema': s['schema'], 'questions': [], 'warnings': ['这是明确的离线演示流程，未调用模型；自有表格请使用模型规划或编辑清洗代码。'], 'mode': 'demo', 'prompt': DEMO_PROMPT}
        s['result'] = None
        return {**public_state(s), 'prompt': DEMO_PROMPT}
    if path == 'upload':
        role = body.get('role')
        if role not in ('source1', 'source2', 'target'):
            raise ContractError('文件角色不支持')
        if body.get('remove') is True:
            if role != 'source2':
                raise ContractError('只能移除可选的源表 B')
            s['tables'].pop(role, None)
            s['raw'].pop(role, None)
        else:
            if body.get('reparse') is True:
                if role not in s['raw']:
                    raise ContractError('尚未上传该文件')
                filename = s['tables'][role]['filename']
                raw = s['raw'][role]
            else:
                filename = Path(str(body.get('filename', ''))).name
                raw = base64.b64decode(body.get('data', ''), validate=True)
            table = read_source(raw, filename, body.get('header_row', 1))
            if role == 'target':
                if any(not r['blank'] for r in table['rows']):
                    raise ContractError('目标模板数据区必须为空；不覆盖已有业务数据')
                s['schema'] = default_schema(table['headers'])
            s['tables'][role] = table
            s['raw'][role] = raw
        s['plan'] = s['result'] = None
        return public_state(s)
    required(s)
    if path == 'plan':
        prompt = body.get('prompt', '')
        if not isinstance(prompt, str) or not prompt.strip() or len(prompt) > 4000:
            raise ContractError('请输入 1 至 4000 字符的处理指令')
        schema = validate_schema(body.get('schema', s['schema']), s['tables']['target']['headers'])
        plan = model_plan(prompt, s['tables'], schema, body.get('share_samples') is True)
        plan['prompt'] = prompt
        s.update(plan=plan, schema=plan['schema'], result=None)
        return public_state(s)
    if path == 'reuse':
        tid = body.get('template_id', '')
        if not isinstance(tid, str) or not re.fullmatch(r'[a-f0-9]{32}', tid):
            raise ContractError('模板编号错误')
        file = template_dir(data) / (tid + '.json')
        if not file.is_file():
            raise ContractError('模板不存在')
        template = json.loads(file.read_text(encoding='utf-8'))
        if template['fingerprint'] != fingerprint(s['tables']):
            raise ContractError('源/目标字段与模板不一致，请重新生成计划；不会猜测复用')
        s['plan'] = {**template['plan'], 'mode': 'template', 'template_id': tid, 'summary': template['summary']}
        s['schema'] = validate_schema(template['schema'], s['tables']['target']['headers'])
        s['result'] = None
        return public_state(s)
    if path == 'preview':
        if body.get('confirmed') is not True:
            raise ContractError('请核对清洗流程与目标约束后勾选确认')
        schema = validate_schema(body.get('schema'), s['tables']['target']['headers'])
        code = body.get('code')
        if s.get('plan') and s['plan'].get('questions') and code == s['plan']['code']:
            raise ContractError('模型仍有未确定问题；请补充指令重新规划，或明确编辑清洗流程')
        planning = s.get('plan') or {'mode': 'manual', 'prompt': body.get('prompt', ''), 'summary': '手动编写受限清洗代码'}
        if code != planning.get('code'):
            planning = {**planning, 'mode': 'manual_edit', 'questions': [], 'code': code, 'original_summary': planning['summary'], 'summary': '人工编辑后的转换流程'}
        if isinstance(body.get('prompt'), str):
            planning = {**planning, 'prompt': body['prompt'][:4000]}
        # A failed new execution must invalidate any previous exportable result.
        s['result'] = None
        result = run_sandbox(s['tables'], code)
        result['issues'] = issues_for(result['rows'], schema, result['source_issues'])
        result['original_preview'] = copy.deepcopy(result['rows'])
        s.update(result=result, schema=schema, plan={**planning, 'code': code, 'schema': schema})
        return public_result(result, schema)
    if path == 'resolve':
        if s['result'] is None:
            raise ContractError('请先生成预览')
        result = apply_decisions(s['result'], body.get('decisions'), s['schema'])
        s['result'] = result
        return public_result(result, s['schema'])
    if path == 'export':
        result = s.get('result')
        if not result or result['issues']:
            raise ContractError('未生成预览或仍有未处理异常，导出已暂停')
        format = body.get('format', 'xlsx')
        encoding = body.get('encoding', 'utf-8-sig')
        raw = export_rows(result['rows'], s['schema'], format, encoding, {**s['tables']['target'], 'raw': s['raw']['target']})
        verify_export(raw, result['rows'], s['schema'], format, encoding)
        run_id = uuid.uuid4().hex
        created_at = datetime.now(timezone.utc).isoformat()
        audit = {
            'run_id': run_id, 'created_at': created_at, 'agent_status': 'completed',
            'plan': s['plan'], 'schema': s['schema'], 'fingerprint': fingerprint(s['tables']),
            'inputs': {k: {'filename': t['filename'], 'sha256': t['sha256'], 'rows': len(t['rows'])} for k, t in s['tables'].items()},
            'format': format, 'encoding': encoding if format == 'csv' else None,
            'output_sha256': hashlib.sha256(raw).hexdigest(), 'output_rows': len(result['rows']),
            'rows': result['rows'], 'before_decisions': result['original_preview'],
            'events': result['events'], 'decisions': result.get('decisions', []),
            'export_readback_verified': True,
        }
        dest = data / run_id
        dest.mkdir(parents=True)
        files = {f'import.{format}': raw, 'audit.json': json.dumps(audit, ensure_ascii=False, indent=2).encode('utf-8'), 'workflow.py': s['plan']['code'].encode('utf-8')}
        for name, content in files.items():
            (dest / name).write_bytes(content)
        buf = io.BytesIO()
        with ZipFile(buf, 'w', ZIP_DEFLATED) as z:
            for name, content in files.items():
                z.writestr(name, content)
        (dest / 'result.zip').write_bytes(buf.getvalue())
        s['last_export'] = audit
        return {'agent_status': 'completed', 'run_id': run_id, 'row_count': len(result['rows']), 'readback_verified': True, 'downloads': {name: f'/download/{run_id}/{name}' for name in [*files, 'result.zip']}}
    if path == 'save':
        audit = s.get('last_export')
        if not audit:
            raise ContractError('请先成功导出一次，再保存该次成功流程')
        name = body.get('name', '').strip()
        if not name or len(name) > 80:
            raise ContractError('模板名称须为 1 至 80 字符')
        tid = uuid.uuid4().hex
        item = {'id': tid, 'name': name, 'created_at': datetime.now(timezone.utc).isoformat(), 'summary': audit['plan']['summary'], 'source_run': audit['run_id'], 'fingerprint': audit['fingerprint'], 'plan': audit['plan'], 'schema': audit['schema']}
        # One-off human cell corrections never become hardcoded workflow values.
        path = template_dir(data) / (tid + '.json')
        path.write_text(json.dumps(item, ensure_ascii=False, indent=2), encoding='utf-8')
        return {'saved': tid, 'templates': templates(data)}
    raise ContractError('通用转换接口不存在')
