import base64
import copy
import io
import json
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app
from core import ContractError
from sandbox import Runtime, SandboxError, validate_code
from universal import *
from universal_api import handle, state


def fixture():
    names = {'source1': '系统A_员工报销.xlsx', 'source2': '系统B_员工费用.csv', 'target': '目标HR_员工导入.xlsx'}
    return {role: read_source((ROOT / 'samples' / 'universal' / name).read_bytes(), name) for role, name in names.items()}


def demo_result():
    result = run_sandbox(fixture(), DEMO_CODE)
    result['issues'] = issues_for(result['rows'], DEMO_SCHEMA, result['source_issues'])
    return result


def corrected():
    r = demo_result()
    issue = next(i for i in r['issues'] if i['column'] == 'Mobile')
    return apply_decisions(r, [{'action': 'edit', 'row_id': issue['row_id'], 'column': 'Mobile', 'value': '13800138001'}], DEMO_SCHEMA)


class UniversalTests(unittest.TestCase):
    def test_two_schemas_cleaning_and_provenance(self):
        r = demo_result()
        self.assertEqual(len(r['rows']), 5)
        self.assertEqual(len(r['issues']), 1)
        self.assertEqual(r['rows'][0]['Emp_ID'], '0001')
        self.assertEqual(r['rows'][0]['Emp_Name'], '张琳')
        self.assertEqual(r['rows'][0]['JoinDate'], '20260115')
        self.assertEqual(r['rows'][0]['Province'], '浙江省')
        self.assertEqual(r['rows'][0]['District'], '西湖区')
        self.assertEqual(r['rows'][0]['Amount'], '1250.50')
        self.assertEqual(r['rows'][0]['_refs'][0]['row'], 2)
        self.assertEqual(len(r['events']), 2)
        self.assertEqual(public_result(r, DEMO_SCHEMA)['agent_status'], 'awaiting_human_confirmation')

    def test_file_perturbation_changes_result(self):
        t = fixture()
        t['source1']['rows'][0]['values']['报销金额'] = '98.25'
        r = run_sandbox(t, DEMO_CODE)
        self.assertEqual(r['rows'][0]['Amount'], '98.25')

    def test_pause_correct_and_revalidate(self):
        r = corrected()
        self.assertFalse(r['issues'])
        self.assertEqual(len(r['decisions']), 1)
        self.assertEqual(public_result(r, DEMO_SCHEMA)['agent_status'], 'ready_to_export')
        again = apply_decisions(r, [{'action': 'edit', 'row_id': '1', 'column': 'JoinDate', 'value': '20260230'}], DEMO_SCHEMA)
        self.assertTrue(again['issues'])

    def test_skip_is_explicit_and_audited(self):
        r = demo_result()
        rid = r['issues'][0]['row_id']
        r = apply_decisions(r, [{'action': 'skip', 'row_id': rid}], DEMO_SCHEMA)
        self.assertEqual(len(r['rows']), 4)
        self.assertFalse(r['issues'])
        self.assertEqual(r['decisions'][0]['action'], 'skip')

    def test_real_xls_roundtrip(self):
        import xlrd
        rows = corrected()['rows']
        data = export_rows(rows, DEMO_SCHEMA, 'xls')
        self.assertTrue(data.startswith(bytes.fromhex('d0cf11e0a1b11ae1')))
        ws = xlrd.open_workbook(file_contents=data).sheet_by_index(0)
        self.assertEqual(ws.cell_value(1, 0), '0001')
        self.assertEqual(ws.cell_type(1, 0), xlrd.XL_CELL_TEXT)
        self.assertEqual(ws.cell_value(1, 4), 1250.5)
        verify_export(data, rows, DEMO_SCHEMA, 'xls', 'utf-8-sig')
        readback = read_source(data, 'old.xls')
        self.assertEqual(readback['rows'][0]['values']['Emp_ID'], '0001')

    def test_xls_formulas_are_not_trusted(self):
        import xlwt
        wb = xlwt.Workbook(); ws = wb.add_sheet('S')
        ws.write(0, 0, '金额'); ws.write(1, 0, xlwt.Formula('1+2'))
        out = io.BytesIO(); wb.save(out)
        with self.assertRaises(ContractError):
            read_source(out.getvalue(), 'formula.xls')

    def test_xlsx_and_all_csv_encodings_roundtrip(self):
        rows = corrected()['rows']
        raw = export_rows(rows, DEMO_SCHEMA)
        verify_export(raw, rows, DEMO_SCHEMA, 'xlsx', 'utf-8-sig')
        ws = load_workbook(io.BytesIO(raw)).active
        self.assertEqual(ws['A2'].data_type, 's')
        self.assertEqual(ws['C2'].value, '20260115')
        for enc in ENCODINGS:
            with self.subTest(encoding=enc):
                raw = export_rows(rows, DEMO_SCHEMA, 'csv', enc)
                verify_export(raw, rows, DEMO_SCHEMA, 'csv', enc)

    def test_csv_unrepresentable_character_fails_without_loss(self):
        rows = corrected()['rows']; rows[0]['Emp_Name'] = '测试😀'
        with self.assertRaises(ContractError): export_rows(rows, DEMO_SCHEMA, 'csv', 'gbk')
        verify_export(export_rows(rows, DEMO_SCHEMA, 'csv', 'utf-8'), rows, DEMO_SCHEMA, 'csv', 'utf-8')

    def test_no_formula_execution_in_exports(self):
        rows = corrected()['rows']; rows[0]['Emp_Name'] = '=HYPERLINK("bad")'
        ws = load_workbook(io.BytesIO(export_rows(rows, DEMO_SCHEMA))).active
        self.assertEqual(ws['B2'].data_type, 's')
        csv_raw = export_rows(rows, DEMO_SCHEMA, 'csv', 'utf-8')
        self.assertIn("'=HYPERLINK", csv_raw.decode())
        verify_export(csv_raw, rows, DEMO_SCHEMA, 'csv', 'utf-8')

    def test_duplicate_conflict_pauses_all_conflicting_outputs(self):
        t = fixture(); t['source2']['rows'][0]['values']['claim_total'] = '999'
        result = run_sandbox(t, DEMO_CODE)
        issues = issues_for(result['rows'], DEMO_SCHEMA)
        self.assertEqual(sum('唯一键冲突' in i['error'] for i in issues), 2)
        rid = next(r['_id'] for r in result['rows'] if r['Amount'] == '999')
        r = apply_decisions(result, [{'action': 'skip', 'row_id': rid}], DEMO_SCHEMA)
        self.assertEqual(sum('唯一键冲突' in i['error'] for i in r['issues']), 0)

    def test_numeric_id_is_not_guessed(self):
        t = fixture(); t['source1']['rows'][0]['values']['员工编号'] = 1
        r = run_sandbox(t, DEMO_CODE)
        self.assertTrue(any('前导零' in i['error'] for i in issues_for(r['rows'], DEMO_SCHEMA)))

    def test_source_formulas_pause_even_when_not_in_output(self):
        t = fixture(); t['source1']['rows'][0]['formula'] = True
        r = run_sandbox(t, DEMO_CODE)
        self.assertEqual(len(r['source_issues']), 1)
        source_issue = r['source_issues'][0]
        resolved = apply_decisions(r, [{'action': 'skip', 'row_id': None, 'issue_id': source_issue['id']}], DEMO_SCHEMA)
        self.assertFalse(resolved['source_issues'])

    def test_invalid_filter_cannot_be_overridden_by_cell_edit(self):
        t = fixture(); t['source1']['rows'][0]['values']['报销金额'] = '不确定'
        r = run_sandbox(t, DEMO_CODE)
        rid = r['rows'][0]['_id']
        self.assertTrue(any(i['skip_only'] for i in issues_for(r['rows'], DEMO_SCHEMA)))
        with self.assertRaises(ContractError):
            apply_decisions(r, [{'action': 'edit', 'row_id': rid, 'column': 'Amount', 'value': '-3'}], DEMO_SCHEMA)

    def test_no_silent_corruption_of_malformed_thousands(self):
        with self.assertRaises(ValueError): numeric('1,2')
        with self.assertRaises(ValueError): numeric('1234567890123456')

    def test_missing_address_is_not_guessed(self):
        t = fixture(); t['source1']['rows'][0]['values']['家庭地址'] = '西湖区某路'
        r = run_sandbox(t, DEMO_CODE)
        self.assertEqual(r['rows'][0]['Province'], '')
        self.assertTrue(any('地址无法' in i['error'] for i in issues_for(r['rows'], DEMO_SCHEMA)))

    def test_restricted_code_rejects_escape_attempts(self):
        attacks = [
            'import os\nresult = table("source1")',
            'result = __import__("os")', 'result = table("source1").__class__',
            'result = open("/etc/passwd")', 'result = table("source1")[0]',
            'result = [x for x in table("source1")]', 'while True: pass',
            'table = 1\nresult = table', 'result = merge([table("source1")]) * 100000000',
            'result = (lambda: 1)()', 'result = table(**{})',
        ]
        for code in attacks:
            with self.subTest(code=code), self.assertRaises(SandboxError): validate_code(code)

    def test_unknown_headers_and_join_are_dynamic(self):
        tables = {'source1': [{ 'X-人员索引': '009', '未知费用字段': '11.25', '_refs': [{'source':'source1','row':2}]}], 'source2': [{'key_code':'009','ORG-label':'研发部','_refs':[{'source':'source2','row':2}]}]}
        code = '''a = map_fields(table("source1"), {"ID":"X-人员索引", "Fee":"未知费用字段"})
b = table("source2")
rows = join(a,b,"ID","key_code",{"Dept":"ORG-label"})
result = select_columns(rows,["ID","Fee","Dept"])'''
        result = Runtime(tables).run(code)
        self.assertEqual(result['rows'][0]['Dept'], '研发部')
        self.assertEqual(len(result['rows'][0]['_refs']), 2)

    def test_schema_cannot_reorder_columns_or_add_regex_code(self):
        s = copy.deepcopy(DEMO_SCHEMA); s['columns'].reverse()
        with self.assertRaises(ContractError): validate_schema(s, fixture()['target']['headers'])
        s = copy.deepcopy(DEMO_SCHEMA); s['columns'][0]['pattern'] = '(a+)+$'
        with self.assertRaises(ContractError): validate_schema(s, fixture()['target']['headers'])

    def test_formula_header_rejected_and_plain_csv_header_remains_text(self):
        wb = Workbook(); wb.active['A1'] = '=1+2'
        buf = io.BytesIO(); wb.save(buf)
        with self.assertRaises(ContractError): read_source(buf.getvalue(),'header.xlsx')
        schema = {'columns':[{'name':'=1+2','type':'text'}],'unique':[]}
        rows = [{'=1+2':'plain','_refs':[],'_id':'1'}]
        exported = export_rows(rows,schema)
        self.assertEqual(load_workbook(io.BytesIO(exported)).active['A1'].data_type,'s')

    def test_diff_tracks_original_values_per_output_row(self):
        r = demo_result()
        p = public_result(r,DEMO_SCHEMA)['preview'][0]
        self.assertEqual(p['before']['Emp_Name'],' 张琳 ')
        self.assertEqual(p['before']['JoinDate'],'2026/01/15')
        self.assertEqual(p['values']['Emp_Name'],'张琳')

    def test_empty_join_key_cannot_match_missing_identifier(self):
        runtime = Runtime({'source1':[{'id':'','_refs':[]}],'source2':[{'id':'','name':'未知员工','_refs':[]}]})
        r = runtime.run('a = table("source1")\nb = table("source2")\nresult = join(a,b,"id","id",{"name":"name"})')
        self.assertEqual(r['rows'][0]['name'],'')
        self.assertTrue(r['rows'][0]['_errors'])

    def test_model_only_gets_headers_by_default(self):
        response = {'summary':'测试', 'code':DEMO_CODE, 'schema':DEMO_SCHEMA, 'questions':[], 'warnings':[]}
        class Fake:
            def __enter__(self): return self
            def __exit__(self,*args): pass
            def read(self,n): return json.dumps({'choices':[{'message':{'content':json.dumps(response)}}]}).encode()
        with patch('universal.config',return_value={'configured':True,'base_url':'https://example.test/v1','model':'test'}), patch.dict(os.environ,{'BRIDGE_API_KEY':'test'}), patch('urllib.request.urlopen',return_value=Fake()) as urlopen:
            model_plan(DEMO_PROMPT, fixture(), DEMO_SCHEMA)
            sent = json.loads(urlopen.call_args.args[0].data)
            content = sent['messages'][1]['content']
            self.assertNotIn('张琳',content); self.assertNotIn('138-0013-8000',content)
            self.assertIn('person_code',content)


class UniversalStateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.data = Path(self.temp.name); self.session = {}
        handle('sample',{},self.session,self.data)

    def tearDown(self): self.temp.cleanup()

    def call(self, path, **body): return handle(path,body,self.session,self.data)

    def preview(self):
        return self.call('preview',code=DEMO_CODE,schema=DEMO_SCHEMA,confirmed=True)

    def finish(self):
        preview = self.preview(); i = preview['issues'][0]
        self.call('resolve',decisions=[{'action':'edit','row_id':i['row_id'],'column':'Mobile','value':'13800138001'}])
        return self.call('export',format='xls')

    def test_export_gate_and_snapshot(self):
        self.preview()
        with self.assertRaises(ContractError): self.call('export',format='xls')
        out = self.finish(); first = self.data / out['run_id'] / 'import.xls'; before = first.read_bytes()
        self.assertTrue(out['readback_verified'])
        self.finish(); self.assertEqual(first.read_bytes(),before)

    def test_templates_require_success_and_recompute(self):
        with self.assertRaises(ContractError): self.call('save',name='先成功再保存')
        self.finish(); saved = self.call('save',name='双系统 HR')
        self.call('sample'); reused = self.call('reuse',template_id=saved['saved'])
        self.assertEqual(reused['plan']['mode'],'template')
        self.assertEqual(self.preview()['issue_count'],1)  # manual phone fix wasn't baked in
        self.session['universal']['tables']['source1']['headers'][0] = 'changed'
        with self.assertRaises(ContractError): self.call('reuse',template_id=saved['saved'])

    def test_failed_new_run_invalidates_old_result(self):
        self.finish()
        with self.assertRaises(ContractError): self.call('preview',code='result = open("secret")',schema=DEMO_SCHEMA,confirmed=True)
        with self.assertRaises(ContractError): self.call('export')

    def test_template_with_business_data_rejected(self):
        wb = Workbook(); wb.active.append(['id']); wb.active.append(['old'])
        raw = io.BytesIO(); wb.save(raw)
        with self.assertRaises(ContractError): self.call('upload',role='target',filename='old.xlsx',data=base64.b64encode(raw.getvalue()).decode())


class UniversalApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(); cls.old_data = app.DATA; app.DATA = Path(cls.temp.name)
        cls.server = app.ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'
        threading.Thread(target=cls.server.serve_forever,daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown(); cls.server.server_close(); app.DATA = cls.old_data; cls.temp.cleanup()

    def post(self,path,body):
        req = urllib.request.Request(self.base+'/api/'+path,data=json.dumps(body).encode(),headers={'Content-Type':'application/json'})
        return json.load(urllib.request.urlopen(req))

    def test_complete_http_human_confirmation_export_and_memory(self):
        sid = self.post('session',{})['session_id']
        def call(path,**body): return self.post('universal/'+path,{'session_id':sid,**body})
        sample = call('sample'); self.assertEqual(sample['tables']['source1']['row_count'],4)
        preview = call('preview',code=DEMO_CODE,schema=DEMO_SCHEMA,confirmed=True)
        self.assertEqual(preview['agent_status'],'awaiting_human_confirmation')
        with self.assertRaises(urllib.error.HTTPError): call('export',format='xls')
        issue = preview['issues'][0]
        ready = call('resolve',decisions=[{'row_id':issue['row_id'],'column':'Mobile','action':'edit','value':'13800138001'}])
        self.assertEqual(ready['agent_status'],'ready_to_export')
        out = call('export',format='xls')
        for name,url in out['downloads'].items():
            self.assertTrue(urllib.request.urlopen(self.base+url).read())
        save = call('save',name='HTTP 成功转换')
        self.assertTrue(save['saved'])
        self.assertIn('跨系统表格转换智能体',urllib.request.urlopen(self.base).read().decode())


if __name__ == '__main__': unittest.main()
