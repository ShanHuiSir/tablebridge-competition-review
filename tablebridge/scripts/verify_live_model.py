"""Opt-in live acceptance: two model calls, synthetic schemas only, no credentials in report."""
from pathlib import Path
import sys,json,copy
from datetime import datetime,timezone
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from universal import read_source,model_plan,run_sandbox,issues_for,apply_decisions,export_rows,verify_export,DEMO_SCHEMA,DEMO_PROMPT
from core import ContractError
PROMPT=DEMO_PROMPT+' 业务口径已确认：同工号且完整内容一致才去重，内容冲突保留并请求人工；异常手机号不推断，按11位手机号约束暂停导出；源日期含 YYYY/MM/DD、YYYY-MM-DD、YYYYMMDD 和 YYYY年M月D日，cast 的 date format 参数指定输出格式，内置解析支持这些输入；地址只拆明确行政后缀，缺失保留异常并人工处理。已确认的这些口径无需再次提问，其他真正不确定之处仍须提问。'
def fixture():
 names={'source1':'系统A_员工报销.xlsx','source2':'系统B_员工费用.csv','target':'目标HR_员工导入.xlsx'}
 return {r:read_source((ROOT/'samples/universal'/name).read_bytes(),name) for r,name in names.items()}
def check(tables,plan):
 assert not plan['questions'],plan['questions']
 result=run_sandbox(tables,plan['code']);result['issues']=issues_for(result['rows'],plan['schema'],result['source_issues']);result['original_preview']=copy.deepcopy(result['rows'])
 assert len(result['rows'])==5 and len(result['issues'])==1 and len(result['events'])==2
 issue=result['issues'][0];assert issue['column']=='Mobile'
 fixed=apply_decisions(result,[{'row_id':issue['row_id'],'issue_id':issue['id'],'column':'Mobile','action':'edit','value':'13800138001'}],plan['schema']);assert not fixed['issues']
 formats={}
 for fmt,enc in [('xlsx','utf-8-sig'),('xls','utf-8-sig'),('csv','gb18030')]:
  raw=export_rows(fixed['rows'],plan['schema'],fmt,enc);verify_export(raw,fixed['rows'],plan['schema'],fmt,enc);formats[fmt]={'readback_verified':True,'bytes':len(raw)}
 first=fixed['rows'][0];assert first['Emp_ID']=='0001' and first['JoinDate']=='20260115' and str(first['Amount'])=='1250.50'
 changed=copy.deepcopy(tables);changed['source1']['rows'][0]['values']['报销金额' if '报销金额' in changed['source1']['headers'] else 'reimbursement_X4']='98.25'
 again=run_sandbox(changed,plan['code']);assert str(again['rows'][0]['Amount'])=='98.25'
 return {'output_rows':5,'issues_before':1,'issues_after':0,'removal_events':2,'first_row':{k:str(first[k]) for k in ['Emp_ID','JoinDate','Amount','Province','City','District']},'formats':formats,'amount_perturbation_verified':'1250.50 → 98.25'}
def main():
    report={'verified_at':datetime.now(timezone.utc).isoformat(),'provider':'DeepSeek','request_model':'deepseek-chat','sample_values_shared':False,'confirmed_prompt':PROMPT,'cases':[]}
    for name,unknown in [('双系统员工转换',False),('变化表头动态规划',True)]:
     tables=fixture()
     if unknown:
      remaps={'source1':dict(zip(tables['source1']['headers'],['employee_ref_Q7','person_label_P8','start_date_Z6','cellphone_N5','reimbursement_X4','residence_L3'])),'source2':dict(zip(tables['source2']['headers'],['staff_ref_U2','name_label_V3','hire_date_W4','phone_contact_X5','expense_amount_Y6','postal_address_Z7']))}
      for role,m in remaps.items():
       t=tables[role];t['headers']=[m[h] for h in t['headers']]
       for row in t['rows']:row['values']={m[k]:v for k,v in row['values'].items()}
     try:
      plan=model_plan(PROMPT,tables,DEMO_SCHEMA);evidence=check(tables,plan)
      report['cases'].append({'case':name,'status':'passed','source_headers':{r:t['headers'] for r,t in tables.items() if r!='target'},'plan':plan,'evidence':evidence})
      print(name,'passed',plan['model'],plan['usage'].get('total_tokens'),flush=True)
     except Exception as e:
      report['cases'].append({'case':name,'status':'failed','error':str(e) if isinstance(e,(ContractError,AssertionError)) else type(e).__name__});print(name,'failed',report['cases'][-1]['error'],flush=True)
     dest=ROOT/'docs/通用Agent模型实测.json';dest.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    report['status']='verified' if all(c['status']=='passed' for c in report['cases']) else 'partial'
    dest.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return 0 if report['status']=='verified' else 1

if __name__ == '__main__':
    raise SystemExit(main())
