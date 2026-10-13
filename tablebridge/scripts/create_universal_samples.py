"""Reproducible, entirely synthetic acceptance fixtures."""
import csv
import io
import sys
from pathlib import Path
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from universal import DEMO_SCHEMA


def create():
    dest = ROOT / 'samples' / 'universal'
    dest.mkdir(parents=True, exist_ok=True)
    a = Workbook()
    ws = a.active
    ws.title = 'OA员工费用'
    ws.append(['员工编号', '员工姓名', '入职日期', '联系电话', '报销金额', '家庭地址'])
    for row in [
        ['0001', ' 张琳 ', '2026/01/15', '138-0013-8000', '1,250.50', '浙江省杭州市西湖区文一路1号'],
        ['0002', '王芳', '2026-01-20', '13900138000', '300', '上海市浦东新区世纪大道2号'],
        ['0003', '陈晨', '2026-01-21', '13600138000', '-20', '广东省深圳市南山区科技路3号'],
        ['0004', '周岚', '2026-01-25', '138AB138000', '28.50', '北京市海淀区中关村大街4号'],
    ]:
        ws.append(row)
    for column in 'ABCDEF':
        ws.column_dimensions[column].width = 23
    a.save(dest / '系统A_员工报销.xlsx')
    out = io.StringIO(newline='')
    writer = csv.writer(out, lineterminator='\r\n')
    writer.writerow(['person_code', 'full_name', 'hired_on', 'mobile', 'claim_total', 'home_address'])
    writer.writerows([
        ['0002', '王芳', '20260120', '13900138000', '300.00', '上海市浦东新区世纪大道2号'],
        ['0005', '李明', '2026年2月1日', '137 0013 8000', '10.25', '江苏省南京市鼓楼区中山路5号'],
        ['0006', '赵宁', '2026-02-02', '13500138000', '60', '广东省广州市天河区体育东路6号'],
    ])
    (dest / '系统B_员工费用.csv').write_bytes(out.getvalue().encode('gbk'))
    target = Workbook()
    target.active.title = 'HR导入'
    target.active.append([c['name'] for c in DEMO_SCHEMA['columns']])
    for column in 'ABCDEFGH':
        target.active.column_dimensions[column].width = 19
    target.save(dest / '目标HR_员工导入.xlsx')
    (dest / '目标约束.json').write_text(__import__('json').dumps(DEMO_SCHEMA, ensure_ascii=False, indent=2), encoding='utf-8')


if __name__ == '__main__':
    create()
