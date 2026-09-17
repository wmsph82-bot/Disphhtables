<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>نموذج إدخال جداول الصيادلة</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --navy:#123C5C;
    --navy-deep:#0C2C44;
    --teal:#0E7C86;
    --teal-light:#E4F2F1;
    --bg:#F5F7F8;
    --card:#FFFFFF;
    --border:#DEE5E9;
    --ink:#1D2B33;
    --muted:#64757F;
    --amber:#9A6B00;
    --amber-bg:#FBF1DC;
    --danger:#B3401F;
    --radius:8px;
    font-family:'Cairo', Tahoma, Arial, sans-serif;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background:var(--bg);
    color:var(--ink);
    font-family:'Cairo', Tahoma, Arial, sans-serif;
    line-height:1.6;
    padding-bottom:60px;
  }
  header.top{
    background:linear-gradient(135deg, var(--navy) 0%, var(--navy-deep) 100%);
    color:#fff;
    padding:26px 20px 30px;
  }
  header.top .wrap{
    max-width:960px;
    margin:0 auto;
  }
  header.top h1{
    margin:0 0 4px;
    font-size:22px;
    font-weight:800;
    letter-spacing:0.2px;
  }
  header.top p.sub{
    margin:0;
    color:#BFD4E0;
    font-size:13.5px;
    font-weight:500;
  }
  .container{
    max-width:960px;
    margin:-16px auto 0;
    padding:0 20px;
  }
  .card{
    background:var(--card);
    border:1px solid var(--border);
    border-radius:var(--radius);
    padding:20px 22px;
    margin-bottom:18px;
  }
  .card h2{
    font-size:15.5px;
    font-weight:700;
    margin:0 0 16px;
    color:var(--navy);
    display:flex;
    align-items:center;
    gap:8px;
  }
  .card h2 .dot{
    width:7px;height:7px;border-radius:50%;background:var(--teal);display:inline-block;
  }
  .meta-card{
    background:var(--navy-deep);
    color:#fff;
    border:none;
    margin-top:0;
    box-shadow:0 6px 18px rgba(12,44,68,0.25);
  }
  .meta-grid{
    display:grid;
    grid-template-columns:repeat(4, 1fr);
    gap:14px;
  }
  @media(max-width:760px){ .meta-grid{grid-template-columns:1fr 1fr;} }
  .field label{
    display:block;
    font-size:12.5px;
    font-weight:600;
    margin-bottom:6px;
    color:var(--muted);
  }
  .meta-card .field label{ color:#AFC7D6; }
  .field input[type=text],
  .field input[type=date],
  .field input[type=time],
  .field input[type=number],
  .field select,
  .field textarea{
    width:100%;
    padding:9px 10px;
    border-radius:6px;
    border:1px solid var(--border);
    font-family:inherit;
    font-size:13.5px;
    background:#fff;
    color:var(--ink);
  }
  .meta-card .field input,
  .meta-card .field select{
    border:1px solid rgba(255,255,255,0.25);
    background:rgba(255,255,255,0.08);
    color:#fff;
  }
  .meta-card .field input::placeholder{ color:#9FB6C4; }
  .meta-card .field input[readonly]{
    background:rgba(255,255,255,0.14);
    font-weight:600;
  }
  .field select option{ color:#000; }
  .row2{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }
  .row3{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; }
  @media(max-width:760px){ .row2,.row3{ grid-template-columns:1fr; } }
  .field{ margin-bottom:14px; }

  .days-grid{
    display:grid;
    grid-template-columns:repeat(2, 1fr);
    gap:8px 18px;
  }
  @media(max-width:600px){ .days-grid{grid-template-columns:1fr;} }
  .day-row{
    display:flex;
    align-items:center;
    gap:10px;
    padding:8px 10px;
    border:1px solid var(--border);
    border-radius:6px;
    background:#FAFBFC;
  }
  .day-row.active{
    border-color:var(--teal);
    background:var(--teal-light);
  }
  .day-row label.chk{
    display:flex;
    align-items:center;
    gap:8px;
    font-size:13.5px;
    font-weight:600;
    min-width:78px;
    cursor:pointer;
  }
  .day-row input[type=checkbox]{
    width:16px;height:16px;accent-color:var(--teal);cursor:pointer;
  }
  .day-row input[type=time]{
    flex:1;
    padding:6px 8px;
    font-size:13px;
    border:1px solid var(--border);
    border-radius:5px;
  }
  .day-row input[type=time]:disabled{
    background:#EEF1F3;
    color:#AAB4BA;
  }

  .rest-days-grid{
    display:flex;
    flex-wrap:wrap;
    gap:8px;
  }
  .rest-chip{
    display:flex;
    align-items:center;
    gap:7px;
    padding:8px 12px;
    border:1px solid var(--border);
    border-radius:20px;
    background:#FAFBFC;
    font-size:13px;
    font-weight:600;
    cursor:pointer;
    user-select:none;
  }
  .rest-chip.active{
    border-color:var(--danger);
    background:#FBEDE7;
    color:var(--danger);
  }
  .rest-chip input{
    width:15px;height:15px;accent-color:var(--danger);cursor:pointer;
  }

  .other-work{ margin-top:8px; display:none; }
  .other-work.show{ display:block; }

  .btn{
    display:inline-flex;
    align-items:center;
    gap:6px;
    padding:10px 18px;
    border-radius:6px;
    border:none;
    font-family:inherit;
    font-size:13.5px;
    font-weight:700;
    cursor:pointer;
  }
  .btn-teal{ background:var(--teal); color:#fff; }
  .btn-teal:hover{ background:#0A6970; }
  .btn-navy{ background:var(--navy); color:#fff; }
  .btn-navy:hover{ background:var(--navy-deep); }
  .btn-outline{ background:#fff; border:1px solid var(--border); color:var(--ink); }
  .btn-outline:hover{ background:#F0F3F4; }
  .btn-danger{ background:#fff; border:1px solid #E9C3B6; color:var(--danger); padding:6px 10px; font-size:12.5px; }
  .btn-danger:hover{ background:#FBEDE7; }

  table{ width:100%; border-collapse:collapse; font-size:12.8px; }
  thead th{
    background:var(--navy);
    color:#fff;
    padding:9px 8px;
    text-align:center;
    font-weight:700;
    font-size:12.3px;
  }
  tbody td{
    padding:8px;
    text-align:center;
    border-bottom:1px solid var(--border);
    vertical-align:middle;
  }
  tbody tr:nth-child(even){ background:#FAFBFC; }
  tbody tr:hover{ background:var(--teal-light); }
  .empty-note{
    text-align:center;
    color:var(--muted);
    padding:22px 10px;
    font-size:13px;
  }
  .table-wrap{ overflow-x:auto; }

  .note-banner{
    background:var(--amber-bg);
    border:1px solid #EBD9AE;
    color:var(--amber);
    padding:10px 14px;
    border-radius:6px;
    font-size:12.5px;
    font-weight:600;
    margin-bottom:16px;
  }

  .actions-row{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    margin-top:10px;
  }

  .sheets-list{ list-style:none; padding:0; margin:0; }
  .sheets-list li{
    display:flex;
    justify-content:space-between;
    align-items:center;
    padding:10px 12px;
    border:1px solid var(--border);
    border-radius:6px;
    margin-bottom:8px;
    font-size:13px;
  }
  .sheets-list .info b{ color:var(--navy); }
  .sheets-list .count{
    background:var(--teal-light);
    color:var(--teal);
    padding:2px 9px;
    border-radius:20px;
    font-size:11.5px;
    font-weight:700;
    margin-inline-start:8px;
  }
  .toast{
    position:fixed;
    bottom:22px;
    left:50%;
    transform:translateX(-50%) translateY(20px);
    background:var(--navy-deep);
    color:#fff;
    padding:11px 20px;
    border-radius:8px;
    font-size:13.5px;
    font-weight:600;
    opacity:0;
    pointer-events:none;
    transition:opacity .25s, transform .25s;
    z-index:50;
    box-shadow:0 8px 22px rgba(0,0,0,0.2);
  }
  .toast.show{ opacity:1; transform:translateX(-50%) translateY(0); }
  .footer-note{
    text-align:center;
    color:var(--muted);
    font-size:11.5px;
    margin-top:24px;
  }
  textarea{ resize:vertical; min-height:70px; }
  .required::after{ content:" *"; color:var(--danger); }
</style>
</head>
<body>

<header class="top">
  <div class="wrap">
    <h1>نموذج إدخال جداول عمل الصيادلة</h1>
    <p class="sub">إدخال بيانات كل صيدلي/ة، ثم حفظ الجدول وتصديره كملف إكسل مُجمّع حسب الصيدلية</p>
  </div>
</header>

<div class="container">

  <!-- Sheet-level meta -->
  <div class="card meta-card">
    <h2 style="color:#fff;"><span class="dot" style="background:#fff;"></span> بيانات الجدول</h2>
    <div class="meta-grid">
      <div class="field">
        <label>تاريخ الإدخال</label>
        <input type="text" id="autoDate" readonly>
      </div>
      <div class="field">
        <label class="required">اسم الصيدلية</label>
        <select id="pharmacySelect"></select>
      </div>
      <div class="field">
        <label>فترة الجدول من</label>
        <input type="date" id="periodFrom">
      </div>
      <div class="field">
        <label>فترة الجدول إلى</label>
        <input type="date" id="periodTo">
      </div>
    </div>
  </div>

  <!-- Staff entry form -->
  <div class="card">
    <h2><span class="dot"></span> إضافة صيدلي / صيدلية</h2>

    <div class="row3">
      <div class="field">
        <label class="required">الاسم</label>
        <input type="text" id="staffName" placeholder="الاسم بالكامل">
      </div>
      <div class="field">
        <label>الرقم الوظيفي (ID)</label>
        <input type="text" id="staffId" placeholder="مثال: 110001234">
      </div>
      <div class="field">
        <label>دفعة التخرج</label>
        <input type="text" id="staffBatch" placeholder="مثال: 2023">
      </div>
    </div>

    <div class="row2">
      <div class="field">
        <label>طبيعة العمل</label>
        <select id="workNature">
          <option value="لونجات">لونجات</option>
          <option value="بارت تايم">بارت تايم</option>
          <option value="دوام كامل (5 أيام)">دوام كامل (5 أيام)</option>
          <option value="أخرى">أخرى</option>
        </select>
        <div class="other-work" id="otherWorkBox">
          <input type="text" id="workNatureOther" placeholder="اكتب طبيعة العمل">
        </div>
      </div>
      <div class="field">
        <label>أيام الراحة (يمكن اختيار أكثر من يوم)</label>
        <div class="rest-days-grid" id="restDaysGrid"></div>
      </div>
    </div>

    <div class="field">
      <label>أيام العمل ومواعيد الحضور (اختر يوم أو أكثر، وحدد الموعد جنب كل يوم)</label>
      <div class="days-grid" id="daysGrid"></div>
    </div>

    <div class="field">
      <label>ملاحظات</label>
      <textarea id="staffNotes" placeholder="أي ملاحظات خاصة بهذا الصيدلي/ة (مثل: ساعة رضاعة)"></textarea>
    </div>

    <button class="btn btn-teal" onclick="addStaffRow()">+ إضافة إلى الجدول</button>
  </div>

  <!-- Current staff table -->
  <div class="card">
    <h2><span class="dot"></span> الجدول الحالي <span id="currentCount" class="count" style="margin-inline-start:6px;"></span></h2>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>م</th><th>الاسم</th><th>ID</th><th>الدفعة</th><th>طبيعة العمل</th>
            <th>أيام العمل / المواعيد</th><th>يوم الراحة</th><th>ملاحظات</th><th></th>
          </tr>
        </thead>
        <tbody id="staffTableBody"></tbody>
      </table>
      <div class="empty-note" id="emptyNote">لسه محدّش اتضاف للجدول — استخدمي النموذج اللي فوق</div>
    </div>
  </div>

  <!-- Sheet-level closing fields -->
  <div class="card">
    <h2><span class="dot"></span> بيانات ختامية</h2>
    <div class="field">
      <label>اسم السنيور المسؤول</label>
      <input type="text" id="seniorName" placeholder="اسم السنيور">
    </div>
    <div class="field">
      <label>ملاحظات أخرى على الجدول</label>
      <textarea id="generalNotes" style="min-height:90px;" placeholder="أي ملاحظات عامة على جدول هذه الصيدلية"></textarea>
    </div>
    <div class="actions-row">
      <button class="btn btn-navy" onclick="saveSheet()">💾 حفظ هذا الجدول</button>
      <button class="btn btn-outline" onclick="resetForm(true)">تفريغ النموذج الحالي</button>
    </div>
  </div>

  <!-- Saved sheets -->
  <div class="card">
    <h2><span class="dot"></span> الجداول المحفوظة <span id="savedCount" class="count"></span></h2>
    <div class="note-banner">
      كل جدول تحفظيه بيتخزن هنا في المتصفح <b>وبيتبعت تلقائيًا</b> لشيت Google Sheets بتاعك. لو الإنترنت مقطوع أو الإرسال فشل، استخدمي زرار "إعادة إرسال" جنب أي جدول، أو صدّري إلى إكسل كنسخة احتياطية.
    </div>
    <ul class="sheets-list" id="sheetsList"></ul>
    <div class="actions-row">
      <button class="btn btn-teal" onclick="exportToExcel()">⭳ تصدير كل الجداول المحفوظة إلى Excel</button>
      <button class="btn btn-navy" onclick="sendAllToGoogle()">☁ إعادة إرسال كل الجداول للشيت</button>
      <button class="btn btn-outline" onclick="clearAllSheets()">مسح كل الجداول المحفوظة</button>
    </div>
  </div>

  <p class="footer-note">التصدير يجمع الجداول تلقائيًا في ملف إكسل واحد، بشيت مستقل لكل صيدلية حسب الاختيار من القائمة أعلاه. الحفظ يرسل نفس التجميع مباشرة لملف Google Sheets المرتبط.</p>
</div>

<div class="toast" id="toast"></div>

<script>
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzxUtT28pbW2J8RcGz17zxGAqBfSwjtINgrddN1B_gaJF79W01I81PIBcruQS1blj7frA/exec";

const PHARMACIES = [
  "صيدلية تأمين خارجي",
  "صيدلية بطاقة خارجي",
  "صيدلية أطفال خارجي",
  "صيدلية الألم تأمين خارجي",
  "صيدلية الألم مجاني خارجي",
  "صيدلية باطني داخلي مجاني (الباب 1 والباب 2 معاً)",
  "صيدلية باطني داخلي تأمين 1",
  "صيدلية باطني داخلي تأمين 2",
  "صيدلية باطني داخلي عمليات 1",
  "صيدلية باطني داخلي عمليات 2"
];
const DAYS = ["السبت","الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة"];

let currentStaff = [];
let savedSheets = [];

function toast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'), 2200);
}

function loadStorage(){
  try{
    const raw = localStorage.getItem('pharmSchedules_v1');
    savedSheets = raw ? JSON.parse(raw) : [];
  }catch(e){ savedSheets = []; }
}
function persistStorage(){
  try{
    localStorage.setItem('pharmSchedules_v1', JSON.stringify(savedSheets));
  }catch(e){ /* storage unavailable — data stays in-memory for this session only */ }
}

function init(){
  // auto date
  const d = new Date();
  document.getElementById('autoDate').value = d.toLocaleDateString('ar-EG', {year:'numeric', month:'long', day:'numeric'});

  // pharmacy select
  const sel = document.getElementById('pharmacySelect');
  PHARMACIES.forEach(p=>{
    const o = document.createElement('option');
    o.value = p; o.textContent = p;
    sel.appendChild(o);
  });

  // rest days (multi-select chips)
  const restGrid = document.getElementById('restDaysGrid');
  DAYS.forEach(day=>{
    const chip = document.createElement('label');
    chip.className = 'rest-chip';
    chip.innerHTML = `<input type="checkbox" class="rest-chk" value="${day}"> ${day}`;
    const chk = chip.querySelector('.rest-chk');
    chk.addEventListener('change', ()=> chip.classList.toggle('active', chk.checked));
    restGrid.appendChild(chip);
  });

  // days grid with time inputs
  const grid = document.getElementById('daysGrid');
  DAYS.forEach(day=>{
    const row = document.createElement('div');
    row.className = 'day-row';
    row.innerHTML = `
      <label class="chk">
        <input type="checkbox" class="day-chk" value="${day}">
        ${day}
      </label>
      <input type="time" class="day-time" disabled>
    `;
    const chk = row.querySelector('.day-chk');
    const time = row.querySelector('.day-time');
    chk.addEventListener('change', ()=>{
      time.disabled = !chk.checked;
      row.classList.toggle('active', chk.checked);
      if(!chk.checked) time.value = '';
    });
    grid.appendChild(row);
  });

  // "other" work nature toggle
  document.getElementById('workNature').addEventListener('change', (e)=>{
    document.getElementById('otherWorkBox').classList.toggle('show', e.target.value === 'أخرى');
  });

  loadStorage();
  renderSavedSheets();
  renderCurrentTable();
}

function addStaffRow(){
  const name = document.getElementById('staffName').value.trim();
  if(!name){ toast('من فضلك اكتبي الاسم'); return; }

  const dayRows = Array.from(document.querySelectorAll('.day-row'));
  const daysData = dayRows
    .filter(r => r.querySelector('.day-chk').checked)
    .map(r => {
      const day = r.querySelector('.day-chk').value;
      const time = r.querySelector('.day-time').value;
      return time ? `${day} (${time})` : day;
    });

  let nature = document.getElementById('workNature').value;
  if(nature === 'أخرى'){
    const other = document.getElementById('workNatureOther').value.trim();
    if(other) nature = other;
  }

  const restDays = Array.from(document.querySelectorAll('.rest-chk'))
    .filter(c => c.checked)
    .map(c => c.value);

  const entry = {
    name,
    id: document.getElementById('staffId').value.trim(),
    batch: document.getElementById('staffBatch').value.trim(),
    nature,
    days: daysData.join(' - '),
    restDay: restDays.join(' - '),
    notes: document.getElementById('staffNotes').value.trim()
  };

  currentStaff.push(entry);
  renderCurrentTable();
  clearStaffFields();
  toast('تمت إضافة ' + name);
}

function clearStaffFields(){
  document.getElementById('staffName').value = '';
  document.getElementById('staffId').value = '';
  document.getElementById('staffBatch').value = '';
  document.getElementById('workNature').value = 'لونجات';
  document.getElementById('workNatureOther').value = '';
  document.getElementById('otherWorkBox').classList.remove('show');
  document.querySelectorAll('.rest-chk').forEach(c=>{
    c.checked = false;
    c.closest('.rest-chip').classList.remove('active');
  });
  document.getElementById('staffNotes').value = '';
  document.querySelectorAll('.day-chk').forEach(c=>{
    c.checked = false;
    c.closest('.day-row').classList.remove('active');
    const t = c.closest('.day-row').querySelector('.day-time');
    t.disabled = true; t.value = '';
  });
  document.getElementById('staffName').focus();
}

function renderCurrentTable(){
  const body = document.getElementById('staffTableBody');
  const emptyNote = document.getElementById('emptyNote');
  body.innerHTML = '';
  document.getElementById('currentCount').textContent = currentStaff.length;

  if(currentStaff.length === 0){
    emptyNote.style.display = 'block';
    return;
  }
  emptyNote.style.display = 'none';

  currentStaff.forEach((s, i)=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${i+1}</td>
      <td>${escapeHtml(s.name)}</td>
      <td>${escapeHtml(s.id)}</td>
      <td>${escapeHtml(s.batch)}</td>
      <td>${escapeHtml(s.nature)}</td>
      <td style="text-align:right;">${escapeHtml(s.days) || '-'}</td>
      <td>${escapeHtml(s.restDay) || '-'}</td>
      <td style="text-align:right;">${escapeHtml(s.notes) || '-'}</td>
      <td><button class="btn btn-danger" onclick="removeStaffRow(${i})">حذف</button></td>
    `;
    body.appendChild(tr);
  });
}

function removeStaffRow(i){
  currentStaff.splice(i,1);
  renderCurrentTable();
}

function escapeHtml(str){
  if(!str) return '';
  return String(str).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}

function saveSheet(){
  if(currentStaff.length === 0){
    toast('ضيفي صيدلي واحد على الأقل قبل الحفظ');
    return;
  }
  const sheet = {
    id: Date.now(),
    savedDate: document.getElementById('autoDate').value,
    pharmacy: document.getElementById('pharmacySelect').value,
    periodFrom: document.getElementById('periodFrom').value,
    periodTo: document.getElementById('periodTo').value,
    senior: document.getElementById('seniorName').value.trim(),
    generalNotes: document.getElementById('generalNotes').value.trim(),
    staff: currentStaff,
    sent: false
  };
  savedSheets.push(sheet);
  persistStorage();
  renderSavedSheets();
  resetForm(false);
  toast('تم حفظ جدول: ' + sheet.pharmacy + ' — جاري الإرسال للشيت...');
  sendSheetToGoogle(sheet);
}

function sendSheetToGoogle(sheet){
  if(!APPS_SCRIPT_URL){ return; }
  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    mode: 'no-cors',
    body: JSON.stringify(sheet)
  }).then(()=>{
    // no-cors responses are opaque — can't confirm success, so mark as "sent" optimistically
    sheet.sent = true;
    persistStorage();
    renderSavedSheets();
    toast('تم إرسال جدول ' + sheet.pharmacy + ' للشيت (افحصيه للتأكد)');
  }).catch(()=>{
    sheet.sent = false;
    persistStorage();
    renderSavedSheets();
    toast('فشل إرسال ' + sheet.pharmacy + ' — تحققي من الإنترنت وأعيدي الإرسال');
  });
}

function sendAllToGoogle(){
  if(savedSheets.length === 0){
    toast('مفيش جداول محفوظة للإرسال');
    return;
  }
  savedSheets.forEach(sheet => sendSheetToGoogle(sheet));
  toast('جاري إرسال كل الجداول للشيت...');
}

function resetForm(showToast){
  currentStaff = [];
  clearStaffFields();
  document.getElementById('periodFrom').value = '';
  document.getElementById('periodTo').value = '';
  document.getElementById('seniorName').value = '';
  document.getElementById('generalNotes').value = '';
  renderCurrentTable();
  if(showToast) toast('تم تفريغ النموذج');
}

function renderSavedSheets(){
  const list = document.getElementById('sheetsList');
  list.innerHTML = '';
  document.getElementById('savedCount').textContent = savedSheets.length;

  if(savedSheets.length === 0){
    list.innerHTML = '<li style="justify-content:center;color:var(--muted);">مفيش جداول محفوظة لسه</li>';
    return;
  }

  savedSheets.slice().reverse().forEach(sheet=>{
    const li = document.createElement('li');
    const period = (sheet.periodFrom || sheet.periodTo)
      ? `${sheet.periodFrom || '—'} → ${sheet.periodTo || '—'}`
      : 'بدون فترة محددة';
    const statusBadge = sheet.sent
      ? '<span class="count" style="background:#E3F3E8;color:#1F7A3D;">اترسل ☁</span>'
      : '<span class="count" style="background:#FBEDE7;color:var(--danger);">لسه ما اترسلش</span>';
    li.innerHTML = `
      <span class="info"><b>${escapeHtml(sheet.pharmacy)}</b> — ${escapeHtml(period)}
        <span class="count">${sheet.staff.length} صيدلي</span>
        ${statusBadge}
      </span>
      <span style="display:flex; gap:6px;">
        <button class="btn btn-outline" style="padding:6px 10px; font-size:12.5px;" onclick="sendSheetToGoogle(savedSheets.find(s=>s.id===${sheet.id}))">إعادة إرسال</button>
        <button class="btn btn-danger" onclick="deleteSheet(${sheet.id})">حذف</button>
      </span>
    `;
    list.appendChild(li);
  });
}

function deleteSheet(id){
  savedSheets = savedSheets.filter(s => s.id !== id);
  persistStorage();
  renderSavedSheets();
  toast('تم حذف الجدول');
}

function clearAllSheets(){
  if(savedSheets.length === 0) return;
  if(!confirm('هل أنتِ متأكدة من مسح كل الجداول المحفوظة؟ لا يمكن التراجع عن هذا.')) return;
  savedSheets = [];
  persistStorage();
  renderSavedSheets();
  toast('تم مسح كل الجداول');
}

function exportToExcel(){
  if(savedSheets.length === 0){
    toast('مفيش جداول محفوظة للتصدير');
    return;
  }

  const wb = XLSX.utils.book_new();
  const grouped = {};
  const allRows = [];
  savedSheets.forEach(sheet=>{
    if(!grouped[sheet.pharmacy]) grouped[sheet.pharmacy] = [];
    sheet.staff.forEach(s=>{
      const row = {
        "اسم الصيدلية": sheet.pharmacy,
        "تاريخ الإدخال": sheet.savedDate,
        "الفترة من": sheet.periodFrom,
        "الفترة إلى": sheet.periodTo,
        "الاسم": s.name,
        "الرقم الوظيفي": s.id,
        "دفعة التخرج": s.batch,
        "طبيعة العمل": s.nature,
        "أيام العمل / المواعيد": s.days,
        "يوم الراحة": s.restDay,
        "ملاحظات": s.notes,
        "السنيور المسؤول": sheet.senior,
        "ملاحظات عامة": sheet.generalNotes
      };
      grouped[sheet.pharmacy].push(row);
      allRows.push(row);
    });
  });

  const colWidths = [
    {wch:26},{wch:14},{wch:12},{wch:12},{wch:24},{wch:14},{wch:14},
    {wch:26},{wch:12},{wch:24},{wch:16},{wch:26}
  ];

  // شيت رئيسي واحد يجمع الكل، مع عمود اسم الصيدلية جنب كل صف
  const summaryWs = XLSX.utils.json_to_sheet(allRows);
  summaryWs['!cols'] = colWidths;
  XLSX.utils.book_append_sheet(wb, summaryWs, 'جدول الصيادلة');

  // + شيت مستقل لكل صيدلية، بنفس عمود الاسم برضو
  Object.keys(grouped).forEach(pharmacy=>{
    const ws = XLSX.utils.json_to_sheet(grouped[pharmacy]);
    ws['!cols'] = colWidths;
    let sheetName = pharmacy.replace(/[\\\/\?\*\[\]:]/g,'').slice(0,31);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  const fileName = `جداول_الصيادلة_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
  toast('تم تصدير الملف');
}

init();
</script>

</body>
</html>
