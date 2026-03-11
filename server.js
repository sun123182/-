const express = require('express');
const crypto = require('@wecom/crypto');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

// ---------- 配置（与企业微信、智能表格后台一致）----------
const TOKEN = process.env.WECOM_TOKEN || 'IQCLf5VMl31IenTIoPk6953';
const AES_KEY = process.env.WECOM_ENCODING_AES_KEY || 'x62C4zUbz8kGWunRHkN8m3t9nyDkzO8zELS3AtcWQ7f';
const SHEET_WEBHOOK =
  process.env.SHEET_WEBHOOK_URL ||
  'https://qyapi.weixin.qq.com/cgi-bin/wedoc/smartsheet/webhook?key=1V8oeRYw2EhwSrkhRkY5LkAbJuQRIGO96pMrj7CKbtrsVBkyfdmfhtvgtpoif0YjEKcACNe4ukyDmxgpIHSWH0vkP7wpz5aRbqNJQ51iK8fI';

// 汇总表 Webhook（可选）：用于「按角色合计」的明细，每条消息写一行，表格内用分组汇总得合计
const SUMMARY_WEBHOOK = process.env.SUMMARY_WEBHOOK_URL || '';
// 汇总表字段 ID：在智能表格创建「汇总明细」表并配置 Webhook 后，从示例 schema 里复制 key 填到下面（或用环境变量）
const SUMMARY_FIELDS = {
  role: process.env.SUMMARY_FIELD_ROLE || 'fRole',
  diff: process.env.SUMMARY_FIELD_DIFF || 'fDiff',
  salary: process.env.SUMMARY_FIELD_SALARY || 'fSalary',
  date: process.env.SUMMARY_FIELD_DATE || 'fDate',
};

// ---------- 1~200 级升级所需经验 ----------
const EXP_TABLE = {
  1: 15, 2: 34, 3: 57, 4: 92, 5: 135, 6: 372, 7: 560, 8: 840, 9: 1242, 10: 1716,
  11: 2360, 12: 3216, 13: 4200, 14: 5460, 15: 7050, 16: 8840, 17: 11040, 18: 13716, 19: 16680, 20: 20216,
  21: 24402, 22: 28980, 23: 34320, 24: 40512, 25: 47216, 26: 54900, 27: 63666, 28: 73080, 29: 83720, 30: 95700,
  31: 108480, 32: 122760, 33: 138666, 34: 155540, 35: 174216, 36: 194832, 37: 216600, 38: 240500, 39: 266682, 40: 294216,
  41: 324240, 42: 356916, 43: 391160, 44: 428280, 45: 468450, 46: 510420, 47: 555680, 48: 604416, 49: 655200, 50: 709716,
  51: 748608, 52: 789631, 53: 832902, 54: 878545, 55: 926689, 56: 977471, 57: 1031036, 58: 1087536, 59: 1147132, 60: 1209994,
  61: 1276301, 62: 1346242, 63: 1420016, 64: 1497832, 65: 1579913, 66: 1666492, 67: 1757815, 68: 1854143, 69: 1955750, 70: 2062925,
  71: 2175973, 72: 2295216, 73: 2420993, 74: 2553663, 75: 2693603, 76: 2841212, 77: 2996910, 78: 3161140, 79: 3334370, 80: 3517093,
  81: 3709829, 82: 3913127, 83: 4127566, 84: 4353756, 85: 4592341, 86: 4844001, 87: 5109452, 88: 5389449, 89: 5684790, 90: 5996316,
  91: 6324914, 92: 6671519, 93: 7037118, 94: 7422752, 95: 7829518, 96: 8258575, 97: 8711144, 98: 9188514, 99: 9692044, 100: 10223168,
  101: 10783397, 102: 11374327, 103: 11997640, 104: 12655110, 105: 13348610, 106: 14080113, 107: 14851703, 108: 15665576, 109: 16524049, 110: 17429566,
  111: 18384706, 112: 19392187, 113: 20454878, 114: 21575805, 115: 22758159, 116: 24005306, 117: 25320796, 118: 26708375, 119: 28171993, 120: 29715818,
  121: 31344244, 122: 33061908, 123: 34873700, 124: 36784778, 125: 38800583, 126: 40926854, 127: 43169645, 128: 45535341, 129: 48030677, 130: 50662758,
  131: 53439077, 132: 56367538, 133: 59456479, 134: 62714694, 135: 66151459, 136: 69776558, 137: 73600313, 138: 77633610, 139: 81887931, 140: 86375389,
  141: 91108760, 142: 96101520, 143: 101367883, 144: 106922842, 145: 112782213, 146: 118962678, 147: 125481832, 148: 132358236, 149: 139611467, 150: 147262175,
  151: 155332142, 152: 163844343, 153: 172823012, 154: 182293713, 155: 192283408, 156: 202820538, 157: 213935103, 158: 225658746, 159: 238024845, 160: 251068606,
  161: 264827165, 162: 279339693, 163: 294647508, 164: 310794191, 165: 327825712, 166: 345790561, 167: 364739883, 168: 384727628, 169: 405810702, 170: 428049128,
  171: 451506220, 172: 476248760, 173: 502347192, 174: 529875818, 175: 558913012, 176: 589541445, 177: 621848316, 178: 655925603, 179: 691870326, 180: 729784819,
  181: 769777027, 182: 811960808, 183: 856456260, 184: 903390063, 185: 952895838, 186: 1005114529, 187: 1060194805, 188: 1118293480, 189: 1179575962, 190: 1244216724,
  191: 1312399800, 192: 1384319309, 193: 1460180007, 194: 1540197871, 195: 1624600714, 196: 1713628833, 197: 1807535693, 198: 1906588648, 199: 2011069705, 200: 2121276324,
};

function expForLevel(level) {
  return EXP_TABLE[level] ?? null;
}

// 解析：账号xxx 等级n 开始经验xxx 结束经验xxx（或 结束升级+xxx）
function parseMessage(text, sender) {
  const raw = (text || '').replace(/@财务账号/g, '').replace(/@财务/g, '').trim();
  const account = raw.match(/账号\s*(\S+)/)?.[1];
  const level = parseInt(raw.match(/等级\s*(\d+)/)?.[1], 10);
  const startStr = raw.match(/(?:开始|经验开始|开始经验)\s*(\d+)/)?.[1];
  const endStr = raw.match(/(?:结束|经验结束|结束经验)\s*(\d+|升级\+?\d+)/)?.[1];

  if (!account || !level || !startStr || !endStr) {
    throw new Error('格式需包含：账号xxx 等级n 开始经验n 结束经验n（或结束升级+n）');
  }

  const expStart = parseInt(startStr, 10);
  const upgradeMatch = endStr.match(/升级\+?(\d+)/);
  let expEnd, diff;

  if (upgradeMatch) {
    const extra = parseInt(upgradeMatch[1], 10);
    const need = expForLevel(level);
    if (need == null) throw new Error(`无等级${level}经验数据`);
    diff = Math.round(need / 10000 - expStart + extra);
    expEnd = `升级+${extra}`;
  } else {
    expEnd = parseInt(endStr, 10);
    if (Number.isNaN(expEnd) || expEnd <= expStart) throw new Error('结束经验需大于开始经验');
    diff = expEnd - expStart;
  }

  let salary;
  if (account === 'mao' || account === '米露') salary = (diff / 1440) * 10;
  else if (level < 160) salary = (diff / 1020) * 10;
  else salary = (diff / 1200) * 10;
  salary = Math.round(salary * 100) / 100;

  const now = new Date();
  const weekdays = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
  return {
    date: now.getDate(),
    weekday: weekdays[now.getDay()],
    wechatName: sender,
    roleName: account,
    level,
    expStart,
    expEnd: String(expEnd),
    diff,
    salary,
    note: '',
    photoTime: '正常拍照',
  };
}

async function appendSheet(row) {
  const body = {
    schema: {
      f1PSmO: '日期', f3MlEe: '星期', fAfOBp: '微信名称', fDe867: '角色名称',
      fR7LNY: '等级', fXUN1W: '经验值（开始）', fXifUI: '经验值（结束）',
      fYOy6d: '差值', fcAmkj: '工资', fpsEdy: '备注', fxB4Oz: '文本11',
    },
    add_records: [{
      values: {
        f1PSmO: row.date,
        f3MlEe: [{ text: row.weekday }],
        fAfOBp: row.wechatName,
        fDe867: row.roleName,
        fR7LNY: row.level,
        fXUN1W: row.expStart,
        fXifUI: row.expEnd,
        fYOy6d: row.diff,
        fcAmkj: row.salary,
        fpsEdy: row.note || '',
        fxB4Oz: [{ text: row.photoTime }],
      },
    }],
  };
  const { data } = await axios.post(SHEET_WEBHOOK, body, {
    headers: { 'Content-Type': 'application/json' },
    timeout: 10000,
  });
  console.log('[表格接口返回]', JSON.stringify(data));
  return data;
}

// ---------- 本地数据存一份（与智能表格同步，供网页统计）----------
const DATA_DIR = path.join(__dirname, 'data');
const RECORDS_FILE = path.join(DATA_DIR, 'records.json');

function loadRecords() {
  try {
    if (fs.existsSync(RECORDS_FILE)) {
      const raw = fs.readFileSync(RECORDS_FILE, 'utf8');
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    }
  } catch (e) {
    console.error('[store] load error', e.message);
  }
  return [];
}

function saveRecords(records) {
  try {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.writeFileSync(RECORDS_FILE, JSON.stringify(records, null, 0), 'utf8');
  } catch (e) {
    console.error('[store] save error', e.message);
  }
}

function saveRecord(row) {
  const dateStr = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  const records = loadRecords();
  records.push({
    date: dateStr,
    roleName: row.roleName,
    diff: row.diff,
    salary: row.salary,
    wechatName: row.wechatName,
    level: row.level,
    expStart: row.expStart,
    expEnd: row.expEnd,
  });
  saveRecords(records);
}

/** 写入汇总明细表：每条消息一行（角色、本次差值、本次工资、日期），表格内用「分组汇总」按角色合计 */
async function appendSummarySheet(row) {
  if (!SUMMARY_WEBHOOK) return null;
  const body = {
    schema: {
      [SUMMARY_FIELDS.role]: '角色名称',
      [SUMMARY_FIELDS.diff]: '本次差值',
      [SUMMARY_FIELDS.salary]: '本次工资',
      [SUMMARY_FIELDS.date]: '日期',
    },
    add_records: [{
      values: {
        [SUMMARY_FIELDS.role]: row.roleName,
        [SUMMARY_FIELDS.diff]: row.diff,
        [SUMMARY_FIELDS.salary]: row.salary,
        [SUMMARY_FIELDS.date]: `${row.date}日`,
      },
    }],
  };
  try {
    const { data } = await axios.post(SUMMARY_WEBHOOK, body, {
      headers: { 'Content-Type': 'application/json' },
      timeout: 10000,
    });
    if (data && data.errcode && data.errcode !== 0) console.error('[汇总表]', data.errcode, data.errmsg);
    else console.log('[汇总表] 已写入 角色=', row.roleName, '差值=', row.diff);
    return data;
  } catch (e) {
    console.error('[汇总表] 请求失败', e.message);
    return null;
  }
}

// ---------- 服务 ----------
const app = express();

app.use((req, res, next) => {
  console.log('[请求]', req.method, req.url);
  next();
});

app.get('/test', (req, res) => {
  res.json({ ok: true, msg: 'wechat bot running' });
});

// 统计页：优先读 public/index.html，没有则发内联页面（避免 Render 未带上 public 时报错）
const INDEX_HTML_PATH = path.join(__dirname, 'public', 'index.html');
const INDEX_HTML_INLINE = `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>角色经验统计</title><style>*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"PingFang SC",sans-serif;margin:0;padding:20px;background:#f5f6f8}h1{font-size:1.35rem;color:#1a1a2e}.sub{color:#666;font-size:.9rem}section{background:#fff;border-radius:10px;padding:16px 20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.08)}section h2{font-size:1rem;color:#333;margin:0 0 12px 0}table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:10px 12px;border-bottom:1px solid #eee}th{color:#666;font-weight:600}.num{text-align:right;font-variant-numeric:tabular-nums}.query-row{display:flex;gap:10px;align-items:center;margin-bottom:12px}.query-row input[type=date]{padding:8px 12px;border:1px solid #ddd;border-radius:6px;font-size:1rem}.query-row button{padding:8px 16px;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer}.empty{color:#999;padding:16px 0}.total{font-weight:600;color:#1a1a2e}</style></head><body><h1>角色经验统计</h1><p class="sub">数据与智能表格同步，从今天起按角色汇总差值总和；可查询任意日期当天练了多少经验。</p><section><h2>从今天起 · 各角色差值总和</h2><p class="sub" style="margin-bottom:12px">统计日期 ≥ <span id="sinceDate">-</span></p><table><thead><tr><th>角色名称</th><th class="num">差值总和</th></tr></thead><tbody id="rolesBody"></tbody><tfoot><tr><td class="total">合计</td><td class="num total" id="rolesTotal">0</td></tr></tfoot></table><p id="rolesEmpty" class="empty" style="display:none">暂无数据（今天起还没有记录）</p></section><section><h2>查询某天练了多少经验（差值）</h2><div class="query-row"><input type="date" id="queryDate"/><button type="button" id="queryBtn">查询</button></div><table id="dayTable" style="display:none"><thead><tr><th>角色名称</th><th class="num">当日差值</th></tr></thead><tbody id="dayBody"></tbody><tfoot><tr><td class="total">当日合计</td><td class="num total" id="dayTotal">0</td></tr></tfoot></table><p id="dayEmpty" class="empty" style="display:none">该日暂无记录</p></section><script>const base=window.location.origin;function todayStr(){const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}async function loadRoleStats(){const from=todayStr();document.getElementById('sinceDate').textContent=from;const res=await fetch(base+'/api/stats/roles?from='+encodeURIComponent(from));const data=await res.json();const tbody=document.getElementById('rolesBody');const totalEl=document.getElementById('rolesTotal');const emptyEl=document.getElementById('rolesEmpty');tbody.innerHTML='';const entries=Object.entries(data.roles||{}).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);let total=0;entries.forEach(([name,sum])=>{total+=sum;const tr=document.createElement('tr');tr.innerHTML='<td>'+name.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</td><td class="num">'+sum.toLocaleString()+'</td>';tbody.appendChild(tr)});totalEl.textContent=total.toLocaleString();emptyEl.style.display=entries.length?'none':'block'}async function queryDay(){const date=document.getElementById('queryDate').value;if(!date)return;const res=await fetch(base+'/api/query?date='+encodeURIComponent(date));const data=await res.json();const table=document.getElementById('dayTable');const tbody=document.getElementById('dayBody');const totalEl=document.getElementById('dayTotal');const emptyEl=document.getElementById('dayEmpty');tbody.innerHTML='';const byRole=data.byRole||{};const entries=Object.entries(byRole).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);let total=0;entries.forEach(([name,sum])=>{total+=sum;const tr=document.createElement('tr');tr.innerHTML='<td>'+name.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</td><td class="num">'+sum.toLocaleString()+'</td>';tbody.appendChild(tr)});totalEl.textContent=total.toLocaleString();table.style.display=entries.length?'table':'none';emptyEl.style.display=entries.length?'none':'block'}document.getElementById('queryBtn').addEventListener('click',queryDay);document.getElementById('queryDate').value=todayStr();loadRoleStats();</script></body></html>`;

app.get('/', (req, res) => {
  try {
    if (fs.existsSync(INDEX_HTML_PATH)) {
      return res.type('html').send(fs.readFileSync(INDEX_HTML_PATH, 'utf8'));
    }
  } catch (_) {}
  res.type('html').send(INDEX_HTML_INLINE);
});

app.use(express.static(path.join(__dirname, 'public')));

// 从某日起各角色差值总和（供网页「从今天开始」统计）
app.get('/api/stats/roles', (req, res) => {
  try {
    const from = req.query.from || ''; // YYYY-MM-DD，前端传「今天」
    const records = loadRecords();
    const filtered = from ? records.filter((r) => r.date >= from) : records;
    const byRole = {};
    for (const r of filtered) {
      const name = r.roleName || '未知';
      byRole[name] = (byRole[name] || 0) + (r.diff || 0);
    }
    res.json({ since: from || null, roles: byRole });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 查询某天练了多少经验（差值）
app.get('/api/query', (req, res) => {
  try {
    const date = req.query.date || ''; // YYYY-MM-DD
    if (!date) return res.status(400).json({ error: '请传 date=YYYY-MM-DD' });
    const records = loadRecords().filter((r) => r.date === date);
    const byRole = {};
    let total = 0;
    for (const r of records) {
      total += r.diff || 0;
      const name = r.roleName || '未知';
      byRole[name] = (byRole[name] || 0) + (r.diff || 0);
    }
    res.json({ date, totalDiff: total, byRole, records });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 回调只收 JSON body（智能机器人标准格式）
app.use('/callback', express.json({ limit: '2mb' }));

app.get('/callback', (req, res) => {
  try {
    const q = req.query;
    // 官方要求：对请求参数做 Urldecode，否则验证不成功
    const msgSig = q.msg_signature ? decodeURIComponent(String(q.msg_signature)) : '';
    const ts = q.timestamp ? decodeURIComponent(String(q.timestamp)) : '';
    const nonce = q.nonce ? decodeURIComponent(String(q.nonce)) : '';
    let echostr = q.echostr ? decodeURIComponent(String(q.echostr)) : '';
    if (!msgSig || !ts || !nonce || !echostr) {
      return res.status(400).send('missing params');
    }
    const sig = crypto.getSignature(TOKEN, ts, nonce, echostr);
    if (sig !== msgSig) {
      console.error('[GET] 签名不符', {
        tokenLen: TOKEN.length,
        tokenPrefix: TOKEN.slice(0, 4) + '...',
        echostrLen: echostr.length,
        computedSig: sig,
        receivedSig: msgSig,
      });
      return res.status(401).send('bad signature');
    }
    const { message } = crypto.decrypt(AES_KEY, echostr);
    console.log('[GET] 验证通过, echostr明文:', message);
    // 响应必须为明文，不能加引号、不能带 BOM、不能带换行符
    res.set('Content-Type', 'text/plain; charset=utf-8');
    res.set('Cache-Control', 'no-cache');
    return res.end(message, 'utf8');
  } catch (e) {
    console.error('[GET] 错误:', e.message);
    return res.status(500).send('error');
  }
});

app.post('/callback', async (req, res) => {
  try {
    const q = req.query;
    const msgSig = q.msg_signature;
    const ts = q.timestamp;
    const nonce = q.nonce;
    if (!msgSig || !ts || !nonce) {
      return res.status(400).send('missing params');
    }
    const body = req.body;
    const encrypt = body && (body.encrypt || body.Encrypt);
    if (!encrypt) {
      console.error('[POST] 无 encrypt');
      return res.status(400).send('no encrypt');
    }
    const sig = crypto.getSignature(TOKEN, ts, nonce, encrypt);
    if (sig !== msgSig) {
      console.error('[POST] 签名不符');
      return res.status(401).send('bad signature');
    }
    const { message } = crypto.decrypt(AES_KEY, encrypt);
    console.log('[POST] 解密成功');

    const msg = (function () {
      try {
        return JSON.parse(message);
      } catch (_) {
        return null;
      }
    })();
    if (!msg || !msg.text || !msg.text.content) {
      console.log('[POST] 非文本或无 content，跳过');
      return res.send('success');
    }
    const content = msg.text.content;
    const sender = (msg.from && msg.from.userid) || '未知';
    console.log('[POST] 内容:', content, '发件人:', sender);

    const row = parseMessage(content, sender);
    console.log('[POST] 解析结果:', row);
    const sheetRes = await appendSheet(row);
    if (sheetRes && sheetRes.errcode && sheetRes.errcode !== 0) {
      console.error('[POST] 表格写入失败', sheetRes.errcode, sheetRes.errmsg);
    } else {
      console.log('[POST] 表格已写入');
      saveRecord(row);
    }
    await appendSummarySheet(row);
    return res.send('success');
  } catch (e) {
    console.error('[POST] 错误:', e.message);
    return res.status(500).send('error');
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log('服务已启动 端口:', port);
  console.log('回调: /callback  主表Webhook:', SHEET_WEBHOOK ? '已配置' : '未配置');
  if (SUMMARY_WEBHOOK) console.log('汇总表Webhook: 已配置（按角色合计明细）');
});
