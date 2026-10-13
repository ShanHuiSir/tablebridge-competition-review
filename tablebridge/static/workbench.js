'use strict';
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let sid, tables = {}, schema = null, plan = null, result = null, busy = false, sourceView = 'source1', savedRuns = false;
const roleLabels = {source1:'源表 A', source2:'源表 B', target:'目标模板'};
function status(message, error = false) { $('status').textContent = message; $('status').className = message ? (error ? 'error' : 'busy') : ''; }
async function request(path, body) {
  const response = await fetch('/api/' + path, {method: body ? 'POST' : 'GET', headers: body ? {'Content-Type':'application/json'} : {}, body: body ? JSON.stringify({session_id:sid, ...body}) : undefined});
  const data = await response.json();
  if (!response.ok) throw Error(data.error || '请求失败');
  return data;
}
const api = (path, body = {}) => request('universal/' + path, body);
function sync() {
  for (const el of document.querySelectorAll('button,input,textarea,select')) el.disabled = busy || !sid;
  $('generate').disabled = busy || !sid || !tables.source1 || !tables.target || !$('prompt').value.trim();
  $('preview').disabled = busy || !sid || !tables.source1 || !tables.target || !$('code').value.trim() || !$('confirm').checked;
  $('export').disabled = busy || !result || result.issue_count > 0;
  $('reuse').disabled = busy || !$('template-select').value || !tables.source1 || !tables.target;
  $('save').disabled = busy || !savedRuns || !$('template-name').value.trim();
}
async function task(fn, message) {
  busy = true; sync(); if (message) status(message);
  try { await fn(); } catch(e) { status(e.message, true); }
  finally { busy = false; sync(); }
}
function invalidate() {
  result = null; savedRuns = false; $('review').hidden = true; $('export-section').hidden = true;
  $('downloads').replaceChildren(); $('save-flow').hidden = true; $('confirm').checked = false; sync();
}
function renderTable(headers, rows, classes = () => '') {
  return '<div class="table-wrap"><table><thead><tr>' + headers.map(h => `<th>${esc(h)}</th>`).join('') + '</tr></thead><tbody>' + rows.map((r,i) => '<tr>' + r.map((v,j) => `<td class="${classes(i,j)}">${esc(typeof v === 'object' && v !== null ? JSON.stringify(v) : v)}</td>`).join('') + '</tr>').join('') + '</tbody></table></div>';
}
function renderSources() {
  for (const role of Object.keys(roleLabels)) {
    const t = tables[role];
    $('info-' + role).textContent = t ? `${t.filename} · ${t.row_count} 行 · ${t.headers.length} 列` : (role === 'source2' ? '可选' : '尚未上传');
    if (t) $('header-' + role).value = t.header_row;
  }
  $('remove-source2').hidden = !tables.source2;
  $('sources-preview').innerHTML = Object.entries(tables).map(([r,t]) => `<p class="preview-name">${esc(roleLabels[r])} · ${esc(t.filename)}</p>${renderTable(['源行',...t.headers], t.preview.map(x => [x.row, ...t.headers.map(h => x.values[h])]))}`).join('');
}
function schemaValue() {
  if (!schema) throw Error('先上传目标模板');
  return JSON.parse(JSON.stringify(schema));
}
function renderSchema() {
  if (!schema) { $('schema-editor').innerHTML = '<p class="note">先上传目标模板。</p>'; return; }
  $('schema-json').value = JSON.stringify(schema, null, 2);
  $('schema-editor').innerHTML = '<div class="table-wrap"><table class="schema-table"><thead><tr><th>目标列（原序）</th><th>类型</th><th>必填</th><th>唯一键</th><th>格式</th><th>下限</th><th>上限</th></tr></thead><tbody>' + schema.columns.map((c,i) => `<tr data-index="${i}"><td>${esc(c.name)}</td><td><select data-field="type" aria-label="${esc(c.name)} 类型">${[['text','文本'],['decimal','数值'],['integer','整数'],['date','日期']].map(([v,label]) => `<option value="${v}" ${c.type === v ? 'selected' : ''}>${label}</option>`).join('')}</select></td><td><input data-field="required" type="checkbox" aria-label="${esc(c.name)} 必填" ${c.required ? 'checked' : ''}></td><td><input data-field="unique" type="checkbox" aria-label="${esc(c.name)} 唯一键" ${(schema.unique || []).includes(c.name) ? 'checked' : ''}></td><td><select data-field="format" aria-label="${esc(c.name)} 格式">${(c.type === 'date' ? [['%Y-%m-%d','YYYY-MM-DD'],['%Y%m%d','YYYYMMDD'],['%Y/%m/%d','YYYY/MM/DD']] : [['','不限'],['digits','纯数字文本'],['phone','11位手机号'],['email','电子邮箱']]).map(([v,label]) => `<option value="${v}" ${(c.type === 'date' ? (c.format || '%Y-%m-%d') : (c.pattern || '')) === v ? 'selected' : ''}>${label}</option>`).join('')}</select></td><td><input class="bound" data-field="min" aria-label="${esc(c.name)} 下限" value="${esc(c.min ?? '')}"></td><td><input class="bound" data-field="max" aria-label="${esc(c.name)} 上限" value="${esc(c.max ?? '')}"></td></tr>`).join('') + '</tbody></table></div>';
}
$('schema-editor').addEventListener('change', event => {
  const el = event.target, row = el.closest('tr[data-index]'); if (!row) return;
  const c = schema.columns[Number(row.dataset.index)], field = el.dataset.field;
  if (field === 'required') c.required = el.checked;
  else if (field === 'unique') schema.unique = schema.columns.filter((x,i) => i === Number(row.dataset.index) ? el.checked : (schema.unique || []).includes(x.name)).map(x => x.name);
  else if (field === 'format') { if (c.type === 'date') c.format = el.value; else { delete c.pattern; if (el.value) c.pattern = el.value; } }
  else if (field === 'type') { c.type = el.value; delete c.format; delete c.pattern; }
  else if (field === 'min' || field === 'max') { delete c[field]; if (el.value.trim()) c[field] = el.value.trim(); }
  invalidate(); renderSchema();
});
$('apply-schema').addEventListener('click', () => { try { const value = JSON.parse($('schema-json').value); if (!value.columns || !Array.isArray(value.columns)) throw Error('约束须有 columns 数组'); schema = value; invalidate(); renderSchema(); status('约束已更新；执行时将校验目标表头和列序。'); } catch(e) { status('JSON 约束无效：' + e.message, true); } });
function renderPlan() {
  $('plan-summary').hidden = !plan;
  if (!plan) { $('code').value = ''; return; }
  const label = {demo:'离线演示流程 · 未调用模型', model:`模型生成 · ${plan.model || ''}`, template:'复用已保存流程', manual_edit:'人工编辑的清洗流程'}[plan.mode] || '手动流程';
  $('plan-summary').innerHTML = `<div class="plan-mode">${esc(label)}</div><p>${esc(plan.summary)}</p>` + [...(plan.questions || []), ...(plan.warnings || [])].map(s => `<p>${esc(s)}</p>`).join('');
  $('code').value = plan.code || ''; renderSchema();
}
function acceptState(data) { tables = data.tables; schema = data.schema; plan = data.plan; invalidate(); renderSources(); renderPlan(); if (data.prompt) $('prompt').value = data.prompt; sync(); }
for (const role of Object.keys(roleLabels)) {
  async function upload() {
    const file = $(role).files[0]; if (!file) return;
    if (file.size > 10 * 1024 * 1024) { status('单文件不能超过 10 MB', true); return; }
    await task(async () => {
      invalidate();
      const data = await new Promise((resolve,reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result.split(',')[1]); reader.onerror = reject; reader.readAsDataURL(file); });
      acceptState(await api('upload', {role, filename:file.name, data, header_row:Number($('header-' + role).value)}));
      status('文件已载入；请生成或复用转换流程。');
    }, '正在读取并检查文件…');
  }
  $(role).addEventListener('change', upload);
  $('header-' + role).addEventListener('change', () => {
    if ($(role).files.length) upload();
    else if (tables[role]) task(async () => {
      const row = Number($('header-' + role).value);
      invalidate();
      try { acceptState(await api('upload', {role,reparse:true,header_row:row})); status('表头行已重新解析，请重新规划。'); }
      catch(e) { $('header-' + role).value = tables[role].header_row; throw e; }
    }, '重新解析指定表头行…');
  });
}
$('remove-source2').addEventListener('click', () => task(async () => { acceptState(await api('upload', {role:'source2',remove:true})); $('source2').value = ''; status('已移除源表 B，请重新规划。'); }));
$('sample').addEventListener('click', () => task(async () => { for (const r of Object.keys(roleLabels)) $(r).value = ''; acceptState(await api('sample')); $('source-details').open = true; status('已载入两份不同结构的模拟表及演示流程。核对后执行预览。'); }, '载入模拟文件…'));
$('generate').addEventListener('click', () => task(async () => { invalidate(); acceptState(await api('plan', {prompt:$('prompt').value, schema:schemaValue(), share_samples:$('share-samples').checked})); $('schema-details').open = true; status(plan.questions?.length ? '模型需要补充信息，请完善指令后重新生成。' : '模型已生成流程。请核对目标约束和清洗代码。'); }, '模型正在分析字段并生成清洗流程…'));
$('prompt').addEventListener('input', () => { invalidate(); plan = null; renderPlan(); sync(); });
$('code').addEventListener('input', () => { invalidate(); sync(); });
$('confirm').addEventListener('change', sync);
$('preview').addEventListener('click', () => task(async () => { result = null; $('downloads').replaceChildren(); $('save-flow').hidden = true; savedRuns = false; $('review').hidden = true; $('export-section').hidden = true; result = await api('preview', {code:$('code').value, schema:schemaValue(), prompt:$('prompt').value, confirmed:$('confirm').checked}); renderResult(); status(result.issue_count ? '转换已暂停导出，请核对高亮异常。' : '校验通过，可以选择目标格式导出。'); $('review').scrollIntoView({block:'start',behavior:'instant'}); }, '在隔离进程中执行清洗并校验结果…'));
function refText(refs) { return refs.map(r => `${roleLabels[r.source] || r.source}:${r.row}`).join(' / '); }
function comparison() {
  const t = tables[sourceView] || tables.source1;
  const bad = new Set(result.issues.filter(i => i.row_id).map(i => i.row_id + ':' + i.column));
  $('comparison').innerHTML = `<div class="compare-pane"><div class="pane-heading"><b>转换前</b><select id="source-view" aria-label="对比源表">${Object.entries(tables).filter(([r]) => r !== 'target').map(([r,t]) => `<option value="${r}" ${r === sourceView ? 'selected' : ''}>${esc(roleLabels[r])} · ${esc(t.filename)}</option>`).join('')}</select></div>${renderTable(['源行', ...t.headers], t.preview.map(r => [r.row, ...t.headers.map(h => r.values[h])]))}</div><div class="compare-pane"><div class="pane-heading"><b>转换后 · 目标列序</b><span>${result.row_count} 行</span></div>${renderTable(['来源行', ...schema.columns.map(c => c.name)], result.preview.map(r => [refText(r.refs), ...schema.columns.map(c => r.values[c.name])]), (i,j) => j === 0 ? 'row-ref' : bad.has(result.preview[i].row_id + ':' + schema.columns[j-1].name) ? 'issue-cell' : String(result.preview[i].before[schema.columns[j-1].name] ?? '') !== String(result.preview[i].values[schema.columns[j-1].name] ?? '') ? 'changed-cell' : '')}</div>`;
  $('source-view').addEventListener('change', e => { sourceView = e.target.value; comparison(); });
}
function renderResult() {
  $('review').hidden = false; $('export-section').hidden = false;
  $('agent-state').textContent = result.issue_count ? '等待人工决策' : '校验通过'; $('agent-state').className = result.issue_count ? 'paused' : '';
  $('metrics').innerHTML = [[result.row_count,'待导出记录'],[result.issue_count,'待处理问题'],[result.event_count,'筛选 / 去重记录'],[result.decision_count,'已提交人工决策']].map(([value,label]) => `<div class="metric"><strong>${value}</strong><span>${label}</span></div>`).join('');
  comparison(); $('issues-panel').hidden = !result.issue_count;
  $('issues').innerHTML = '<table><thead><tr><th>来源行</th><th>字段</th><th>当前值</th><th>需要核对</th><th>决策</th><th>核实后的值</th></tr></thead><tbody>' + result.issues.map((i,n) => `<tr data-issue="${n}"><td class="row-ref">${esc(refText(i.refs))}</td><td>${esc(i.column)}</td><td class="issue-cell">${esc(typeof i.value === 'object' ? JSON.stringify(i.value) : i.value)}</td><td>${esc(i.error)}</td><td><select data-action aria-label="问题 ${n+1} 决策"><option value="">暂不处理</option>${i.row_id && !result.issues.some(x => x.row_id === i.row_id && x.skip_only) && schema.columns.some(c => c.name === i.column) ? '<option value="edit">修正此值</option>' : ''}<option value="skip">跳过整行</option></select></td><td><input data-value aria-label="问题 ${n+1} 核实后的值" placeholder="核实后填写"></td></tr>`).join('') + '</tbody></table>';
  $('events-summary').textContent = `查看筛选与去重记录（${result.event_count} 条，显示前 100 条）`;
  $('events').innerHTML = renderTable(['来源行','处理理由'], result.events.map(e => [refText(e.refs),e.reason]));
  $('export-note').textContent = result.issue_count ? '存在未处理异常，导出已暂停。提交决策后会重新验证。' : '目标列序及类型校验通过。生成文件后会再次回读验证。';
  sync();
}
$('resolve').addEventListener('click', () => task(async () => { const decisions = []; for (const row of $('issues').querySelectorAll('tr[data-issue]')) { const issue = result.issues[Number(row.dataset.issue)], action = row.querySelector('[data-action]').value; if (action) decisions.push({row_id:issue.row_id, issue_id:issue.id, column:issue.column, action, value:row.querySelector('[data-value]').value}); } if (!decisions.length) throw Error('先为至少一个问题选择决策'); result = await api('resolve',{decisions}); $('downloads').replaceChildren(); $('save-flow').hidden = true; savedRuns = false; renderResult(); status(result.issue_count ? '决策已记录，仍有问题需要处理。' : '全部异常已处理并通过校验，可以导出。'); }, '提交人工决策并重新校验…'));
$('format').addEventListener('change', () => { $('encoding-label').hidden = $('format').value !== 'csv'; });
$('export').addEventListener('click', () => task(async () => { const data = await api('export',{format:$('format').value, encoding:$('encoding').value}); $('downloads').innerHTML = Object.entries(data.downloads).map(([name,url]) => `<a class="download" href="${esc(url)}" download>${esc(name === 'result.zip' ? '下载结果整包' : name)}</a>`).join(''); savedRuns = true; $('save-flow').hidden = false; $('export-note').textContent = `已生成 ${data.row_count} 条记录，回读验证通过。运行编号 ${data.run_id.slice(0,8)}。`; status('导出成功。可下载导入文件、审计记录和本次清洗代码。'); }, '正在生成文件并回读验证…'));
function renderTemplates(items) { $('template-select').innerHTML = '<option value="">选择已保存流程</option>' + items.map(t => `<option value="${esc(t.id)}">${esc(t.name)}</option>`).join(''); sync(); }
$('save').addEventListener('click', () => task(async () => { const data = await api('save',{name:$('template-name').value}); renderTemplates(data.templates); $('template-select').value = data.saved; status('成功流程已保存。下次上传同类表格后，可在左侧复用。'); }));
$('template-name').addEventListener('input', sync); $('template-select').addEventListener('change', sync);
$('reuse').addEventListener('click', () => task(async () => { const data = await api('reuse',{template_id:$('template-select').value}); acceptState(data); $('prompt').value = plan.prompt || ''; status('流程已复用。请确认当前文件和目标约束，再执行预览。'); }));
task(async () => { const session = await request('session',{}); sid = session.session_id; const data = await request('config'); $('model-status').textContent = data.model.configured ? `已配置模型 · ${data.model.model}` : '未配置模型 · 可运行演示'; renderTemplates((await api('templates')).templates); status('上传源表与目标模板开始，或载入双系统演示。'); });
