/* 小郑AI外贸工具站 - 静态版数据分析工具
   说明：全部在浏览器本地完成，不上传服务器。
*/

const state = {
  lastResults: {},
  crm: JSON.parse(localStorage.getItem('xz_crm_records') || '[]')
};

const fieldAliases = {
  productId: ['产品ID','商品ID','产品编号','产品id','Product ID','ProductID','product_id','ID','id','pid','Product No.','Item ID','产品链接ID'],
  title: ['产品名称','商品名称','标题','产品标题','Product Name','Title','Item Name','Product Title'],
  impressions: ['曝光','曝光量','展示','展示量','Impressions','Impression','Exposure','Views','PV'],
  clicks: ['点击','点击量','Clicks','Click','访问点击','产品点击'],
  visitors: ['访客','访客数','访问人数','Visitors','Visitor','UV'],
  inquiries: ['询盘','询盘数','商机','商机数','询盘个数','Inquiries','Inquiry','Business Opportunities','Opportunities','Leads'],
  tm: ['TM','TM咨询','TM咨询数','TradeManager','Trademanager','TM咨询买家数','沟通','咨询','Chat','Chats'],
  orders: ['订单','订单数','成交订单','Orders','Order','Transactions'],
  cost: ['花费','消耗','广告花费','推广花费','Cost','Spend','Ad Spend','费用'],
  sales: ['成交金额','销售额','GMV','Sales','Revenue','Order Amount'],
  country: ['国家','国家/地区','地区','Country','Region','Market'],
  buyerLevel: ['买家等级','客户等级','Buyer Level','Level','买家类型'],
  stay: ['停留时长','访问时长','平均停留时长','Stay Time','Duration','Avg. Visit Duration'],
  pages: ['访问页面数','浏览页面数','页面数','Pages','Page Views','PV/UV'],
  repeatVisits: ['重复访问','重复访问次数','访问次数','Visit Count','Visits','Repeat Visits']
};

function normalizeKey(key) {
  return String(key || '')
    .replace(/\s+/g, '')
    .replace(/[()（）:：/\\_-]/g, '')
    .toLowerCase();
}

function detectFields(rows) {
  const keys = rows.length ? Object.keys(rows[0]) : [];
  const map = {};
  for (const [target, aliases] of Object.entries(fieldAliases)) {
    const aliasSet = aliases.map(normalizeKey);
    const found = keys.find(k => aliasSet.includes(normalizeKey(k)) || aliasSet.some(a => normalizeKey(k).includes(a)));
    map[target] = found || null;
  }
  return map;
}

function num(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/,/g, '').replace(/%/g, '').replace(/[¥$￥]/g, '').trim();
  const parsed = Number(cleaned);
  return isFinite(parsed) ? parsed : 0;
}

function text(value) {
  return String(value ?? '').trim();
}

function pct(n) {
  if (!isFinite(n)) return '0%';
  return `${(n * 100).toFixed(2)}%`;
}

function money(n) {
  if (!isFinite(n)) return '0';
  return Number(n).toFixed(2);
}

function quantile(arr, q) {
  const xs = arr.filter(v => isFinite(v)).sort((a,b) => a-b);
  if (!xs.length) return 0;
  const pos = (xs.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (xs[base + 1] !== undefined) return xs[base] + rest * (xs[base + 1] - xs[base]);
  return xs[base];
}

async function readFile(fileInputId) {
  const input = document.getElementById(fileInputId);
  const file = input?.files?.[0];
  if (!file) throw new Error('请先选择 Excel / CSV 文件');
  if (!window.XLSX) throw new Error('表格解析库未加载成功，请刷新页面或检查网络');

  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  if (!rows.length) throw new Error('未读取到有效数据，请确认表格第一张工作表有数据');
  return { rows, sheetName, fileName: file.name };
}

function renderError(containerId, err) {
  document.getElementById(containerId).innerHTML = `<div class="result-section"><span class="badge danger">读取失败</span><p>${escapeHtml(err.message || err)}</p></div>`;
  document.getElementById(containerId).classList.remove('empty');
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>'"]/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[s]));
}

function badge(label, type='info') {
  return `<span class="badge ${type}">${escapeHtml(label)}</span>`;
}

function buildTable(rows, columns, maxRows=50) {
  if (!rows.length) return '<p>暂无符合条件的数据。</p>';
  const head = columns.map(c => `<th>${escapeHtml(c.label)}</th>`).join('');
  const body = rows.slice(0, maxRows).map(row => `<tr>${columns.map(c => `<td>${escapeHtml(typeof c.value === 'function' ? c.value(row) : row[c.key])}</td>`).join('')}</tr>`).join('');
  return `<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>${rows.length > maxRows ? `<p class="mini-help">仅预览前 ${maxRows} 行，可导出完整CSV。</p>` : ''}`;
}

function summaryCards(items) {
  return `<div class="summary-cards">${items.map(i => `<div class="summary-card"><strong>${escapeHtml(i.value)}</strong><span>${escapeHtml(i.label)}</span></div>`).join('')}</div>`;
}

function fieldsInfo(map) {
  const pairs = Object.entries(map).filter(([,v]) => v).map(([k,v]) => `${k} = ${v}`);
  return pairs.length ? `<div class="mini-help">已识别字段：${escapeHtml(pairs.join('；'))}</div>` : '';
}

function setResult(containerId, html) {
  const el = document.getElementById(containerId);
  el.classList.remove('empty');
  el.innerHTML = html;
}

function downloadCsv(filename, rows) {
  if (!rows || !rows.length) {
    alert('暂无可导出的数据');
    return;
  }
  const headers = Array.from(rows.reduce((set, row) => {
    Object.keys(row).forEach(k => set.add(k));
    return set;
  }, new Set()));
  const csv = [headers.join(',')].concat(rows.map(row => headers.map(h => csvCell(row[h])).join(','))).join('\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function csvCell(v) {
  const s = String(v ?? '');
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function resultActions(key, filename) {
  return `<div class="result-actions"><button class="btn ghost" onclick="downloadCsv('${filename}', state.lastResults['${key}'] || [])">导出CSV结果</button></div>`;
}

async function analyzeProduct() {
  try {
    const { rows, fileName } = await readFile('productFile');
    const map = detectFields(rows);
    const data = rows.map((r, idx) => {
      const impressions = num(r[map.impressions]);
      const clicks = num(r[map.clicks]);
      const inquiries = num(r[map.inquiries]);
      const tm = num(r[map.tm]);
      const visitors = num(r[map.visitors]);
      const leads = inquiries + tm;
      const ctr = impressions ? clicks / impressions : 0;
      const inquiryRate = clicks ? leads / clicks : 0;
      return {
        序号: idx + 1,
        产品ID: text(r[map.productId]) || `第${idx + 1}行`,
        产品标题: text(r[map.title]).slice(0, 80),
        曝光: impressions,
        点击: clicks,
        访客: visitors,
        询盘: inquiries,
        TM: tm,
        商机合计: leads,
        点击率: pct(ctr),
        商机转化率: pct(inquiryRate),
        分类: '',
        优先级: '',
        建议动作: ''
      };
    });
    const medImp = quantile(data.map(d => d.曝光), .5);
    const medClick = quantile(data.map(d => d.点击), .5);
    const medCtr = quantile(data.map(d => parseFloat(d.点击率) / 100), .5);

    data.forEach(d => {
      const ctrRaw = d.曝光 ? d.点击 / d.曝光 : 0;
      if (d.曝光 >= medImp && d.点击 >= medClick && d.商机合计 > 0) {
        d.分类 = '优爆品'; d.优先级 = 'P1'; d.建议动作 = '保留核心资源位，优先优化主图第二张、详情页卖点和RFQ承接，可适当增加广告预算。';
      } else if (d.曝光 >= medImp && ctrRaw < medCtr) {
        d.分类 = '高曝光低点击潜力品'; d.优先级 = 'P1'; d.建议动作 = '优先换主图、标题前30字符、价格展示和核心关键词，观察7-14天。';
      } else if (d.点击 >= medClick && d.商机合计 === 0) {
        d.分类 = '有点击无商机待优化'; d.优先级 = 'P2'; d.建议动作 = '排查详情页、MOQ、价格、证书、定制服务、交期和信任背书，必要时做A/B主图。';
      } else if (d.曝光 < medImp * 0.5 && d.点击 < medClick * 0.5 && d.商机合计 === 0) {
        d.分类 = '低效待观察'; d.优先级 = 'P3'; d.建议动作 = '不建议继续占用重点资源；先补关键词或改为长尾承接品，连续30天无起色可下架/暂停。';
      } else {
        d.分类 = '正常观察'; d.优先级 = 'P2'; d.建议动作 = '保持基础维护，结合利润和客户质量决定是否继续推广。';
      }
    });

    const sorted = data.sort((a,b) => orderPriority(a.优先级) - orderPriority(b.优先级) || b.商机合计 - a.商机合计 || b.点击 - a.点击);
    state.lastResults.product = sorted;
    const counts = countBy(sorted, '分类');
    const html = summaryCards([
      { label: '读取产品数', value: sorted.length },
      { label: '优爆品', value: counts['优爆品'] || 0 },
      { label: '潜力/待优化', value: (counts['高曝光低点击潜力品'] || 0) + (counts['有点击无商机待优化'] || 0) },
      { label: '低效待观察', value: counts['低效待观察'] || 0 }
    ]) + resultActions('product', '产品爆品筛选结果.csv') +
      `<div class="result-section"><h3>产品筛选结果</h3>${fieldsInfo(map)}${buildTable(sorted, [
        {key:'产品ID', label:'产品ID'}, {key:'产品标题', label:'产品标题'}, {key:'曝光', label:'曝光'}, {key:'点击', label:'点击'}, {key:'询盘', label:'询盘'}, {key:'TM', label:'TM'}, {key:'点击率', label:'点击率'}, {key:'商机转化率', label:'商机转化率'}, {key:'分类', label:'分类'}, {key:'优先级', label:'优先级'}, {key:'建议动作', label:'建议动作'}
      ])}</div>` +
      `<div class="result-section"><h3>运营结论</h3><ul class="kv-list"><li>先处理：高曝光低点击产品，最快影响点击率。</li><li>重点检查：有点击无商机产品，通常是详情页、价格、MOQ、信任背书问题。</li><li>资源倾斜：已有询盘/TM的优爆品，适合做广告承接链接。</li><li>暂停观察：低效产品不要长期占用橱窗、活动和广告资源。</li></ul></div>`;
    setResult('productResult', html);
  } catch (err) { renderError('productResult', err); }
}

function orderPriority(p) {
  return {P1: 1, P2: 2, P3: 3}[p] || 9;
}

function countBy(rows, key) {
  return rows.reduce((acc, row) => { acc[row[key]] = (acc[row[key]] || 0) + 1; return acc; }, {});
}

async function analyzeAds() {
  try {
    const { rows } = await readFile('adsFile');
    const map = detectFields(rows);
    const data = rows.map((r, idx) => {
      const impressions = num(r[map.impressions]);
      const clicks = num(r[map.clicks]);
      const cost = num(r[map.cost]);
      const inquiries = num(r[map.inquiries]);
      const tm = num(r[map.tm]);
      const orders = num(r[map.orders]);
      const leads = inquiries + tm;
      return {
        序号: idx + 1,
        产品ID: text(r[map.productId]) || `第${idx + 1}行`,
        产品标题: text(r[map.title]).slice(0, 80),
        国家: text(r[map.country]),
        曝光: impressions,
        点击: clicks,
        花费: cost,
        询盘: inquiries,
        TM: tm,
        订单: orders,
        商机合计: leads,
        CTR: pct(impressions ? clicks / impressions : 0),
        CPC: clicks ? money(cost / clicks) : '0.00',
        CPL: leads ? money(cost / leads) : '无商机',
        诊断: '',
        动作: ''
      };
    });
    const avgCost = data.length ? data.reduce((s,d)=>s+d.花费,0) / data.length : 0;
    const leadRows = data.filter(d => d.商机合计 > 0);
    const avgCpl = leadRows.length ? leadRows.reduce((s,d)=>s+(d.花费/d.商机合计),0) / leadRows.length : 0;

    data.forEach(d => {
      const cpl = d.商机合计 ? d.花费 / d.商机合计 : Infinity;
      if (d.点击 >= 20 && d.商机合计 === 0) {
        d.诊断 = '20点击无商机'; d.动作 = '建议暂停或降预算，优先检查承接链接、主图、价格、MOQ。';
      } else if (d.花费 >= avgCost * 1.5 && d.商机合计 === 0) {
        d.诊断 = '高消耗无转化'; d.动作 = '进入浪费清单，建议暂停观察7天，必要时更换推广品。';
      } else if (d.商机合计 > 0 && (!avgCpl || cpl <= avgCpl)) {
        d.诊断 = '低CPL有效品'; d.动作 = '建议保留加码，作为全站推核心承接链接。';
      } else if (d.商机合计 > 0 && cpl > avgCpl * 1.5) {
        d.诊断 = '有商机但成本偏高'; d.动作 = '保留但控预算，优化详情页询盘转化和客户质量。';
      } else {
        d.诊断 = '正常观察'; d.动作 = '继续观察7-14天，结合国家和买家等级再判断。';
      }
    });

    const countryStats = groupByCountry(data).map(c => {
      let advice = '正常观察';
      if (c.点击 >= 20 && c.商机合计 === 0) advice = '屏蔽候选：点击多但无商机';
      else if (c.商机合计 > 0 && c.花费 / c.商机合计 <= avgCpl) advice = '保留：商机成本较好';
      return { ...c, CTR: pct(c.曝光 ? c.点击 / c.曝光 : 0), CPL: c.商机合计 ? money(c.花费 / c.商机合计) : '无商机', 建议: advice };
    }).sort((a,b)=> b.花费 - a.花费);

    const sorted = data.sort((a,b) => rankAd(a.诊断) - rankAd(b.诊断) || b.花费 - a.花费 || b.点击 - a.点击);
    state.lastResults.ads = sorted;
    state.lastResults.adsCountries = countryStats;
    const counts = countBy(sorted, '诊断');
    const html = summaryCards([
      { label: '读取记录数', value: sorted.length },
      { label: '建议暂停/降预算', value: (counts['20点击无商机'] || 0) + (counts['高消耗无转化'] || 0) },
      { label: '低CPL有效品', value: counts['低CPL有效品'] || 0 },
      { label: '国家记录数', value: countryStats.length }
    ]) + `<div class="result-actions"><button class="btn ghost" onclick="downloadCsv('全站推广告体检结果.csv', state.lastResults.ads || [])">导出产品诊断CSV</button><button class="btn ghost" onclick="downloadCsv('国家屏蔽建议.csv', state.lastResults.adsCountries || [])">导出国家建议CSV</button></div>` +
      `<div class="result-section"><h3>广告产品诊断</h3>${fieldsInfo(map)}${buildTable(sorted, [
        {key:'产品ID', label:'产品ID'}, {key:'产品标题', label:'产品标题'}, {key:'国家', label:'国家'}, {key:'点击', label:'点击'}, {key:'花费', label:'花费'}, {key:'询盘', label:'询盘'}, {key:'TM', label:'TM'}, {key:'CPC', label:'CPC'}, {key:'CPL', label:'CPL'}, {key:'诊断', label:'诊断'}, {key:'动作', label:'动作'}
      ])}</div>` +
      `<div class="result-section"><h3>国家屏蔽/保留建议</h3>${buildTable(countryStats, [
        {key:'国家', label:'国家'}, {key:'曝光', label:'曝光'}, {key:'点击', label:'点击'}, {key:'花费', label:'花费'}, {key:'商机合计', label:'商机'}, {key:'CTR', label:'CTR'}, {key:'CPL', label:'CPL'}, {key:'建议', label:'建议'}
      ])}</div>` +
      `<div class="result-section"><h3>下周动作建议</h3><ul class="kv-list"><li>先暂停：20点击无商机且高花费的产品。</li><li>先加码：低CPL有效品，优先作为主推承接。</li><li>先排查：点击有了但询盘没有的产品详情页。</li><li>先屏蔽：点击多、花费高、无商机的国家。</li></ul></div>`;
    setResult('adsResult', html);
  } catch (err) { renderError('adsResult', err); }
}

function rankAd(label) {
  return {'20点击无商机':1,'高消耗无转化':2,'有商机但成本偏高':3,'低CPL有效品':4,'正常观察':5}[label] || 9;
}

function groupByCountry(data) {
  const map = new Map();
  data.forEach(d => {
    const key = d.国家 || '未识别国家';
    const cur = map.get(key) || { 国家:key, 曝光:0, 点击:0, 花费:0, 商机合计:0 };
    cur.曝光 += d.曝光; cur.点击 += d.点击; cur.花费 += d.花费; cur.商机合计 += d.商机合计;
    map.set(key, cur);
  });
  return Array.from(map.values());
}

function analyzeDashboard() {
  const form = new FormData(document.getElementById('dashboardForm'));
  const m = Object.fromEntries(Array.from(form.entries()).map(([k,v]) => [k, num(v)]));
  const ctr = m.impressions ? m.clicks / m.impressions : 0;
  const visitorRate = m.clicks ? m.visitors / m.clicks : 0;
  const leadRate = m.visitors ? (m.inquiries + m.tm) / m.visitors : 0;
  const inquiryRate = m.visitors ? m.inquiries / m.visitors : 0;
  const cpl = (m.inquiries + m.tm) ? m.cost / (m.inquiries + m.tm) : 0;
  const orderRate = (m.inquiries + m.tm) ? m.orders / (m.inquiries + m.tm) : 0;
  const roi = m.cost ? m.sales / m.cost : 0;
  const rows = [
    {指标:'点击率 CTR', 数值:pct(ctr), 判断: ctr < .015 ? '偏低' : ctr < .03 ? '正常观察' : '较好', 建议: ctr < .015 ? '优先优化主图、标题前30字符、价格展示和产品排序。' : '继续保持，重点提升询盘转化。'},
    {指标:'访客承接率', 数值:pct(visitorRate), 判断: visitorRate < .6 ? '偏低' : '正常', 建议: '检查点击后是否进入有效产品页，关注无线端加载和产品首屏。'},
    {指标:'商机转化率', 数值:pct(leadRate), 判断: leadRate < .03 ? '偏低' : leadRate < .08 ? '正常' : '较好', 建议: leadRate < .03 ? '详情页增加定制能力、证书、MOQ说明、包装案例和询盘引导。' : '可以扩大优质流量。'},
    {指标:'广告CPL', 数值: m.cost ? money(cpl) : '未填写花费', 判断: cpl ? '需结合行业' : '无数据', 建议: '按产品ID拆分看，不要只看账户平均。'},
    {指标:'订单转化率', 数值:pct(orderRate), 判断: orderRate < .03 ? '偏低' : '正常', 建议: orderRate < .03 ? '重点复盘报价、样品、付款、MOQ、交期和客户预算差距。' : '沉淀成交客户案例，做复购。'},
    {指标:'ROI', 数值: roi ? roi.toFixed(2) : '未填写成交金额', 判断: roi >= 3 ? '较好' : roi > 0 ? '待提升' : '无数据', 建议: roi >= 3 ? '可保持投放，优化利润品占比。' : '先控浪费产品，再提高询盘质量。'}
  ];
  state.lastResults.dashboard = rows;
  const problem = rows.filter(r => ['偏低','待提升'].includes(r.判断)).map(r => r.指标).join('、') || '暂无明显异常';
  const html = summaryCards([
    { label: '曝光量', value: m.impressions },
    { label: '点击率', value: pct(ctr) },
    { label: '商机合计', value: m.inquiries + m.tm },
    { label: '订单数', value: m.orders }
  ]) + resultActions('dashboard', '店铺月度经营看板.csv') +
    `<div class="result-section"><h3>月度经营诊断</h3>${buildTable(rows, [
      {key:'指标', label:'指标'}, {key:'数值', label:'数值'}, {key:'判断', label:'判断'}, {key:'建议', label:'建议'}
    ])}</div>` +
    `<div class="result-section"><h3>本月核心结论</h3><p>当前优先排查项：<strong>${escapeHtml(problem)}</strong>。</p><ul class="kv-list"><li>先看流量入口：曝光和点击率是否达标。</li><li>再看承接页面：点击后有没有询盘或TM。</li><li>再看销售跟进：报价后有没有推进到样品/付款。</li><li>最后看复盘周期：建议每7天小调整，30天大复盘。</li></ul></div>`;
  setResult('dashboardResult', html);
}

function clearDashboard() {
  document.getElementById('dashboardForm').reset();
  const el = document.getElementById('dashboardResult');
  el.classList.add('empty');
  el.textContent = '填写核心指标后生成看板。';
}

async function analyzeVisitor() {
  try {
    const { rows } = await readFile('visitorFile');
    const map = detectFields(rows);
    const data = rows.map((r, idx) => {
      const country = text(r[map.country]) || '未识别国家';
      const level = text(r[map.buyerLevel]) || '未识别等级';
      const pages = num(r[map.pages]);
      const stay = num(r[map.stay]);
      const repeat = num(r[map.repeatVisits]);
      const inquiries = num(r[map.inquiries]);
      const tm = num(r[map.tm]);
      let score = 0;
      if (/L[1-9]\+?|金|高|premium|verified/i.test(level)) score += 30;
      if (pages >= 3) score += 20;
      if (stay >= 60) score += 20;
      if (repeat >= 2) score += 15;
      if (inquiries + tm > 0) score += 30;
      let label = score >= 70 ? '高质量访客' : score >= 40 ? '可二次开发' : '普通访客';
      return { 序号: idx + 1, 国家: country, 买家等级: level, 访问页面数: pages, 停留时长: stay, 重复访问: repeat, 询盘: inquiries, TM: tm, 质量分: score, 分类: label, 建议: visitorAdvice(label) };
    }).sort((a,b) => b.质量分 - a.质量分);
    const countryStats = groupVisitorCountry(data);
    state.lastResults.visitor = data;
    state.lastResults.visitorCountries = countryStats;
    const counts = countBy(data, '分类');
    const html = summaryCards([
      { label: '读取访客数', value: data.length },
      { label: '高质量访客', value: counts['高质量访客'] || 0 },
      { label: '可二次开发', value: counts['可二次开发'] || 0 },
      { label: '国家数量', value: countryStats.length }
    ]) + `<div class="result-actions"><button class="btn ghost" onclick="downloadCsv('高质量访客识别结果.csv', state.lastResults.visitor || [])">导出访客CSV</button><button class="btn ghost" onclick="downloadCsv('访客国家分析.csv', state.lastResults.visitorCountries || [])">导出国家CSV</button></div>` +
      `<div class="result-section"><h3>访客质量明细</h3>${fieldsInfo(map)}${buildTable(data, [
        {key:'国家', label:'国家'}, {key:'买家等级', label:'买家等级'}, {key:'访问页面数', label:'访问页面数'}, {key:'停留时长', label:'停留时长'}, {key:'重复访问', label:'重复访问'}, {key:'询盘', label:'询盘'}, {key:'TM', label:'TM'}, {key:'质量分', label:'质量分'}, {key:'分类', label:'分类'}, {key:'建议', label:'建议'}
      ])}</div>` +
      `<div class="result-section"><h3>国家质量排行</h3>${buildTable(countryStats, [
        {key:'国家', label:'国家'}, {key:'访客数', label:'访客数'}, {key:'高质量访客', label:'高质量访客'}, {key:'平均质量分', label:'平均质量分'}, {key:'建议', label:'建议'}
      ])}</div>`;
    setResult('visitorResult', html);
  } catch (err) { renderError('visitorResult', err); }
}

function visitorAdvice(label) {
  if (label === '高质量访客') return '建议当天二次开发，匹配产品目录、案例和轻定制方案。';
  if (label === '可二次开发') return '建议进入观察池，结合国家和访问产品判断是否主动开发。';
  return '暂不投入过多销售精力，可用于判断广告流量质量。';
}

function groupVisitorCountry(data) {
  const map = new Map();
  data.forEach(d => {
    const cur = map.get(d.国家) || { 国家:d.国家, 访客数:0, 高质量访客:0, 分数合计:0 };
    cur.访客数 += 1;
    cur.高质量访客 += d.分类 === '高质量访客' ? 1 : 0;
    cur.分数合计 += d.质量分;
    map.set(d.国家, cur);
  });
  return Array.from(map.values()).map(c => ({
    国家:c.国家,
    访客数:c.访客数,
    高质量访客:c.高质量访客,
    平均质量分: (c.分数合计 / c.访客数).toFixed(1),
    建议: c.高质量访客 >= 2 ? '重点国家，可用于广告保留/加码' : c.访客数 >= 10 && c.高质量访客 === 0 ? '低质流量偏多，广告需观察' : '正常观察'
  })).sort((a,b)=> Number(b.平均质量分) - Number(a.平均质量分));
}

function initCrm() {
  document.getElementById('crmForm').addEventListener('submit', e => {
    e.preventDefault();
    const form = new FormData(e.target);
    const record = Object.fromEntries(form.entries());
    record.createdAt = new Date().toISOString().slice(0,10);
    record.id = Date.now();
    state.crm.unshift(record);
    saveCrm();
    e.target.reset();
    renderCrm();
  });
  renderCrm();
}

function saveCrm() {
  localStorage.setItem('xz_crm_records', JSON.stringify(state.crm));
}

function renderCrm() {
  const today = new Date().toISOString().slice(0,10);
  const rows = state.crm.map(r => ({
    客户名称: r.name,
    国家: r.country,
    需求产品: r.product,
    数量: r.quantity,
    阶段: r.stage,
    下次跟进: r.nextDate,
    状态: r.nextDate && r.nextDate <= today ? '今日/逾期需跟进' : '正常',
    当前卡点: r.note,
    创建日期: r.createdAt
  }));
  state.lastResults.crm = rows;
  const counts = countBy(rows, '阶段');
  const due = rows.filter(r => r.状态 === '今日/逾期需跟进').length;
  const html = summaryCards([
    { label:'客户总数', value: rows.length },
    { label:'今日需跟进', value: due },
    { label:'已报价S3', value: counts['S3 已报价'] || 0 },
    { label:'推进中S4', value: counts['S4 样品/付款推进'] || 0 }
  ]) + `<div class="result-section"><h3>客户跟进清单</h3>${buildTable(rows, [
    {key:'客户名称', label:'客户名称'}, {key:'国家', label:'国家'}, {key:'需求产品', label:'需求产品'}, {key:'数量', label:'数量'}, {key:'阶段', label:'阶段'}, {key:'下次跟进', label:'下次跟进'}, {key:'状态', label:'状态'}, {key:'当前卡点', label:'当前卡点'}
  ], 100)}</div>` +
  `<div class="result-section"><h3>S1-S5跟进提醒</h3><ul class="kv-list"><li>S1：先问客户类型、职位、用途、国家。</li><li>S2：确认数量、预算、时间线、包装和证书。</li><li>S3：报价后24-48小时跟进，给替代方案。</li><li>S4：围绕样品、付款、交期推进订单。</li></ul></div>`;
  setResult('crmResult', html);
}

function exportCrm() { downloadCsv('B2B询盘跟进CRM.csv', state.lastResults.crm || []); }
function clearCrm() {
  if (!confirm('确认清空本地CRM数据？')) return;
  state.crm = [];
  saveCrm();
  renderCrm();
}

function bindActions() {
  document.body.addEventListener('click', e => {
    const action = e.target?.dataset?.action;
    if (!action) return;
    if (action === 'analyze-product') analyzeProduct();
    if (action === 'analyze-ads') analyzeAds();
    if (action === 'analyze-dashboard') analyzeDashboard();
    if (action === 'clear-dashboard') clearDashboard();
    if (action === 'analyze-visitor') analyzeVisitor();
    if (action === 'export-crm') exportCrm();
    if (action === 'clear-crm') clearCrm();
  });
}

window.addEventListener('DOMContentLoaded', () => {
  bindActions();
  initCrm();
});
