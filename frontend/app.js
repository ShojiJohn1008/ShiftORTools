const tableBody = document.getElementById('table-body');
const rowTpl = document.getElementById('row-template');
const messageEl = document.getElementById('message');
const calendarModal = document.getElementById('calendar-modal');
const calTitle = document.getElementById('cal-title');
const calGrid = document.getElementById('cal-grid');
const calPrev = document.getElementById('cal-prev');
const calNext = document.getElementById('cal-next');
const calSave = document.getElementById('cal-save');
const calClose = document.getElementById('cal-close');

let currentCfg = {};
let editingHospital = null;
let editingYear = null;
let editingMonth = null; // 1-12
let scheduleCache = {}; // month -> assignments
let residentHighlight = null; // {name: string, month: 'YYYY-MM', ng_dates: []}
let residentsCache = {}; // month -> residents list
let currentDragMeta = null; // {resident, from_date, from_hospital}
let lastChanges = []; // stack of manual changes for undo

function pushChange(change){
  lastChanges.push(change);
  updateUndoButtonState();
}

function popChange(){
  const c = lastChanges.pop();
  updateUndoButtonState();
  return c;
}

function updateUndoButtonState(){
  const btn = document.getElementById('undo-change');
  if(!btn) return;
  if(lastChanges.length === 0){ btn.disabled = true; btn.classList.add('disabled'); }
  else { btn.disabled = false; btn.classList.remove('disabled'); }
}

async function undoLastChange(){
  if(lastChanges.length === 0){ showMessage('取り消す変更はありません','error'); return; }
  const ch = popChange();
  try{
    if(ch.type === 'assign'){
      const p = {month: ch.data.month, date: ch.data.date, resident: ch.data.resident};
      const res = await fetch('/api/manual_unassign', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(p)});
      if(!res.ok){ const t = await res.text(); showMessage('取り消し失敗: '+t,'error'); return; }
      const jd = await res.json(); if(jd && jd.result) scheduleCache[ch.data.month] = jd.result;
      showMessage('変更を取り消しました','success');
      renderCalendar(); if(document.getElementById('preview-area')) renderScheduleTable(scheduleCache[ch.data.month], 'preview-area'); if(document.getElementById('run-result')) renderScheduleTable(scheduleCache[ch.data.month], 'run-result');
    }else if(ch.type === 'unassign'){
      // re-apply the removed assignment(s). data.hospitals may be an array
      const month = ch.data.month; const date = ch.data.date; const resident = ch.data.resident;
      const hospitals = ch.data.hospitals || [];
      for(const h of hospitals){
        const p = {month: month, date: date, resident: resident, hospital: h, max_assignments: Number(document.getElementById('max-assignments')?.value || 2)};
        const res = await fetch('/api/manual_assign', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(p)});
        if(!res.ok){ const t = await res.text(); showMessage('再割当失敗: '+t,'error'); return; }
        const jd = await res.json(); if(jd && jd.result) scheduleCache[month] = jd.result;
      }
      showMessage('変更を取り消しました（再割当）','success');
      renderCalendar(); if(document.getElementById('preview-area')) renderScheduleTable(scheduleCache[month], 'preview-area'); if(document.getElementById('run-result')) renderScheduleTable(scheduleCache[month], 'run-result');
    }else if(ch.type === 'move'){
      // reverse move: move from to_date back to from_date
      const payload = {month: ch.data.month, resident: ch.data.resident, from_date: ch.data.to_date, from_hospital: ch.data.to_hospital, to_date: ch.data.from_date, to_hospital: ch.data.from_hospital, max_assignments: Number(document.getElementById('max-assignments')?.value || 2)};
      const res = await fetch('/api/manual_move', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
      if(!res.ok){ const t = await res.text(); showMessage('元に戻す失敗: '+t,'error'); return; }
      const jd = await res.json(); if(jd && jd.result) scheduleCache[ch.data.month] = jd.result;
      showMessage('移動を取り消しました','success');
      renderCalendar(); if(document.getElementById('preview-area')) renderScheduleTable(scheduleCache[ch.data.month], 'preview-area'); if(document.getElementById('run-result')) renderScheduleTable(scheduleCache[ch.data.month], 'run-result');
    }
  }catch(e){ showMessage('取り消し中にエラーが発生しました: '+e.message,'error'); }
}

function showMessage(text, level='info'){
  messageEl.textContent = text;
  messageEl.className = level;
  setTimeout(()=>{ messageEl.textContent=''; messageEl.className=''; }, 4000);
}

async function loadConfig(){
  try{
    const res = await fetch('/api/config');
    if(!res.ok) throw new Error(await res.text());
    currentCfg = await res.json();
    renderTable(currentCfg);
    showMessage('設定を読み込みました。');
  }catch(e){
    showMessage('読み込みエラー: '+e.message, 'error');
    currentCfg = {};
    renderTable(currentCfg);
  }
}

function renderTable(cfg){
  tableBody.innerHTML='';
  const names = Object.keys(cfg).sort();
  if(names.length===0){
    addHospitalRow();
    return;
  }
  for(const name of names){
    const row = rowTpl.content.cloneNode(true);
    const tr = row.querySelector('tr');
    tr.querySelector('.h-name-input').value = name;
    // attach original name for rename handling
    tr.dataset.origName = name;
    // show calendar-config summary (count of date keys)
    const summaryEl = tr.querySelector('.cal-summary');
    const dateKeys = Object.keys(cfg[name] || {}).filter(k=>/^\d{4}-\d{2}-\d{2}$/.test(k));
    summaryEl.textContent = `${dateKeys.length} 日`;
    attachRowHandlers(tr);
    // attach calendar button
    tr.querySelector('.calendar').addEventListener('click', ()=> openCalendar(tr.querySelector('.h-name-input').value.trim()));
    tableBody.appendChild(tr);
  }
}

function attachRowHandlers(tr){
  tr.querySelector('.remove').addEventListener('click', ()=>{
    tr.remove();
  });
}

function addHospitalRow(){
  const row = rowTpl.content.cloneNode(true);
  const tr = row.querySelector('tr');
  attachRowHandlers(tr);
  tr.querySelector('.calendar').addEventListener('click', ()=> {
    const nm = tr.querySelector('.h-name-input').value.trim();
    if(!nm){ showMessage('病院名を入力してからカレンダーを開いてください','error'); return; }
    openCalendar(nm);
  });
  tableBody.appendChild(tr);
}

async function saveConfig(){
  // Build new config from table rows; preserve per-date calendar data when renaming
  const rows = Array.from(tableBody.querySelectorAll('tr'));
  const newCfg = {};
  for(const tr of rows){
    const name = tr.querySelector('.h-name-input').value.trim();
    if(!name) continue;
    const orig = tr.dataset.origName || '';
    if(orig && currentCfg[orig]){
      // preserve existing calendar entries under original name
      newCfg[name] = currentCfg[orig];
    }else if(currentCfg[name]){
      newCfg[name] = currentCfg[name];
    }else{
      newCfg[name] = {};
    }
  }
  currentCfg = newCfg;

  try{
    const res = await fetch('/api/config', {
      method: 'PUT',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(currentCfg)
    });
    if(!res.ok) throw new Error(await res.text());
    showMessage('保存しました。', 'success');
    // reload to normalize and refresh table
    await loadConfig();
  }catch(e){
    showMessage('保存エラー: '+e.message, 'error');
  }
}

// Calendar editor
async function openCalendar(hospital){
  editingHospital = hospital || '';
  if(!editingHospital){ showMessage('病院名を指定してください','error'); return; }
  // ensure object exists
  if(!currentCfg[editingHospital]) currentCfg[editingHospital] = {};
  const now = new Date();
  editingYear = now.getFullYear();
  editingMonth = now.getMonth() + 1;
  const monthKey = `${editingYear}-${String(editingMonth).padStart(2,'0')}`;
  await fetchSchedule(monthKey);
  renderCalendar();
  calendarModal.classList.remove('hidden');
  calTitle.textContent = `${editingHospital} — ${editingYear}/${String(editingMonth).padStart(2,'0')}`;
  // ensure undo button exists above calendar grid
  ensureUndoButton();
}

function closeCalendar(){
  calendarModal.classList.add('hidden');
  editingHospital = null;
  residentHighlight = null;
}

function ensureUndoButton(){
  // add an undo button just above calGrid inside the modal if not present
  if(document.getElementById('undo-change')) return;
  const container = calendarModal.querySelector('.cal-header') || calendarModal;
  const btn = document.createElement('button'); btn.id = 'undo-change'; btn.className = 'undo-change'; btn.textContent = '変更をやり直す';
  btn.disabled = true; btn.classList.add('disabled');
  btn.addEventListener('click', ()=> undoLastChange());
  // insert before calGrid
  calendarModal.insertBefore(btn, calGrid);
}

function renderCalendar(){
  calGrid.innerHTML = '';
  const year = editingYear;
  const month = editingMonth;
  const first = new Date(year, month-1, 1);
  const last = new Date(year, month, 0);
  // weekday header
  const header = document.createElement('div'); header.className='cal-row cal-header';
  ['日','月','火','水','木','金','土'].forEach(d=>{ const c=document.createElement('div'); c.className='cal-cell cal-head'; c.textContent=d; header.appendChild(c); });
  calGrid.appendChild(header);

  let row = document.createElement('div'); row.className='cal-row';
  // fill leading blanks
  for(let i=0;i<first.getDay();i++){
    const cell = document.createElement('div'); cell.className='cal-cell empty'; row.appendChild(cell);
  }
  for(let d=1; d<=last.getDate(); d++){
    const dateStr = `${year}-${String(month).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const cell = document.createElement('div'); cell.className='cal-cell day-cell';
    // mark NG days for resident highlight view
    if(residentHighlight && residentHighlight.month === `${year}-${String(month).padStart(2,'0')}`){
      // NG dates from parsed resident data (take precedence)
      const isNg = Array.isArray(residentHighlight.ng_dates) && residentHighlight.ng_dates.indexOf(dateStr) !== -1;
      if(isNg){
        cell.classList.add('ng-day');
      } else {
        // Also shade days where this resident already has an assignment (any hospital) with a different style
        try{
          const monthKey = `${year}-${String(month).padStart(2,'0')}`;
          if(scheduleCache[monthKey] && scheduleCache[monthKey].assignments){
            const dayAssign = scheduleCache[monthKey].assignments[dateStr] || {};
            for(const hosp in dayAssign){
              const arr = dayAssign[hosp] || [];
              if(Array.isArray(arr) && arr.indexOf(residentHighlight.name) !== -1){
                cell.classList.add('assigned-day');
                // show which hospital(s) the resident is assigned to on this date
                try{
                  const names = [];
                  const dayAssignAll = scheduleCache[monthKey] && scheduleCache[monthKey].assignments && scheduleCache[monthKey].assignments[dateStr] ? scheduleCache[monthKey].assignments[dateStr] : {};
                  for(const hh in dayAssignAll){
                    const a = dayAssignAll[hh] || [];
                    if(Array.isArray(a) && a.indexOf(residentHighlight.name) !== -1){ names.push(hh); }
                  }
                  if(names.length){
                    const lab = document.createElement('div'); lab.className = 'assigned-label'; lab.textContent = names.join(', ');
                    // make assigned label draggable so user can drag this resident to another date
                    lab.draggable = true;
                    // store metadata on dragstart: resident name, source date, source hospital (if multiple, first)
                    lab.addEventListener('dragstart', (ev)=>{
                      const meta = {resident: residentHighlight.name, from_date: dateStr, from_hospital: names[0] || null};
                      ev.dataTransfer.setData('application/json', JSON.stringify(meta));
                      try{ ev.dataTransfer.effectAllowed = 'move'; }catch(e){}
                    });
                    cell.appendChild(lab);
                  }
                }catch(e){ }
                break;
              }
            }
          }
        }catch(e){ /* ignore errors */ }
      }
    }
    const label = document.createElement('div'); label.className='cal-day-label'; label.textContent = d;
    const inp = document.createElement('input'); inp.type='number'; inp.min='0'; inp.className='cal-day-input'; inp.value = currentCfg[editingHospital] && currentCfg[editingHospital][dateStr] !== undefined ? currentCfg[editingHospital][dateStr] : 0;
    // when viewing resident NG highlight, make inputs read-only
    const monthKey = `${year}-${String(month).padStart(2,'0')}`;
    if(residentHighlight && residentHighlight.month === monthKey){ inp.disabled = true; }
    cell.appendChild(label); cell.appendChild(inp);
    // assigned names box
    const assignedBox = document.createElement('div'); assignedBox.className='assigned-names';
    if(scheduleCache[monthKey] && scheduleCache[monthKey].assignments){
      const arr = (scheduleCache[monthKey].assignments[dateStr] && scheduleCache[monthKey].assignments[dateStr][editingHospital]) || [];
      if(arr.length){
        const list = document.createElement('div'); list.className='assigned-list'; list.textContent = arr.join(', ');
        assignedBox.appendChild(list);
      }
    }
    cell.appendChild(assignedBox);
    // attach click handler for manual assign when residentHighlight mode
    cell.addEventListener('click', (e)=>{
      if(!residentHighlight) return;
      const monthStr = `${year}-${String(month).padStart(2,'0')}`;
      if(residentHighlight.month !== monthStr){ showMessage('表示中の住民と月が一致しません','error'); return; }
      // don't allow on NG days
      if(Array.isArray(residentHighlight.ng_dates) && residentHighlight.ng_dates.indexOf(dateStr) !== -1){ showMessage('この日はNGです。割当できません。','error'); return; }
      openAssignModal(dateStr, residentHighlight.name, monthStr);
    });
    // drag & drop handlers to accept a resident dragged from another cell
    cell.addEventListener('dragover', (e)=>{
      if(!residentHighlight) return;
      // prevent dropping on NG days (red hatch)
      const isNgCell = cell.classList.contains('ng-day') || (residentHighlight && Array.isArray(residentHighlight.ng_dates) && residentHighlight.ng_dates.indexOf(dateStr) !== -1);
      if(isNgCell){
        // show visual not-allowed; do NOT call preventDefault so drop is disabled
        cell.classList.add('drop-not-allowed');
        return;
      }
      e.preventDefault();
      cell.classList.add('drop-target');
    });
    cell.addEventListener('dragleave', (e)=>{ if(!residentHighlight) return; cell.classList.remove('drop-target'); cell.classList.remove('drop-not-allowed'); });
    cell.addEventListener('drop', async (e)=>{
      if(!residentHighlight) return;
      // if ng cell, ignore drop
      const isNgCell = cell.classList.contains('ng-day') || (residentHighlight && Array.isArray(residentHighlight.ng_dates) && residentHighlight.ng_dates.indexOf(dateStr) !== -1);
      if(isNgCell){ showMessage('NG日にはドロップできません','error'); cell.classList.remove('drop-not-allowed'); return; }
      e.preventDefault();
      cell.classList.remove('drop-target');
      try{
        const raw = e.dataTransfer.getData('application/json');
        if(!raw) return;
        const meta = JSON.parse(raw);
        // don't allow dropping onto same date/hospital
        const targetDate = dateStr;
        if(meta.from_date === targetDate){ showMessage('同じ日への移動です','error'); return; }
        // compute candidate hospitals for target date similar to openAssignModal
        const monthKey = `${year}-${String(month).padStart(2,'0')}`;
        if(!scheduleCache[monthKey]) await fetchSchedule(monthKey);
        const sched = scheduleCache[monthKey] || {};
        let hospitals = [];
        if(sched && sched.hospitals && Array.isArray(sched.hospitals)) hospitals = sched.hospitals.slice();
        if(!hospitals || hospitals.length===0) hospitals = Object.keys(currentCfg || {});
        let isHoliday = false;
        try{ const hres = await fetch(`/api/is_holiday?date=${encodeURIComponent(targetDate)}`); if(hres.ok){ const hj = await hres.json(); isHoliday = !!hj.is_holiday; } }catch(e){}
        const candidates = [];
        for(const hosp of hospitals){
          let capacity = 0;
          try{
            const cfg = currentCfg[hosp] || {};
            if(cfg[targetDate] !== undefined){ capacity = Number(cfg[targetDate]) || 0; }
            else {
              const dt = new Date(targetDate);
              const jsDay = dt.getDay();
              const wk = (jsDay + 6) % 7;
              if(cfg[String(wk)] !== undefined){ capacity = Number(cfg[String(wk)]) || 0; }
            }
          }catch(e){ capacity = 0; }
          const assignedArr = (sched.assignments && sched.assignments[targetDate] && sched.assignments[targetDate][hosp]) || [];
          const assignedCount = Array.isArray(assignedArr) ? assignedArr.length : 0;
          const dt = new Date(targetDate); const isWeekend = dt.getDay() === 0 || dt.getDay() === 6;
          if(hosp === '大学病院' && (isWeekend || isHoliday)) capacity = capacity + 4;
          const remaining = capacity - assignedCount;
          if(remaining > 0 || (Array.isArray(assignedArr) && assignedArr.indexOf(meta.resident) !== -1)){
            candidates.push({hosp, remaining});
          }
        }
        if(candidates.length === 0){ showMessage('移動先に空きスロットがありません','error'); return; }

        // Show a compact modal to choose which hospital to move into
        const overlay2 = document.createElement('div'); overlay2.className = 'assign-modal-overlay';
        const box2 = document.createElement('div'); box2.className = 'assign-modal-box';
        const h2 = document.createElement('h3'); h2.textContent = `${meta.resident} を ${meta.from_date} から ${targetDate} に移動`; box2.appendChild(h2);
        const info2 = document.createElement('div'); info2.textContent = '移動先の病院を選択してください'; box2.appendChild(info2);
        const list2 = document.createElement('div'); list2.className = 'assign-list';
        for(const it of candidates){
          const b = document.createElement('button'); b.className = 'assign-btn'; b.innerHTML = `${it.hosp}<br><small>残${it.remaining}</small>`;
          b.addEventListener('click', async ()=>{
            try{
              const uiMax = Number(document.getElementById('max-assignments')?.value || 0) || 2;
              const payload = {month: monthKey, resident: meta.resident, from_date: meta.from_date, from_hospital: meta.from_hospital, to_date: targetDate, to_hospital: it.hosp, max_assignments: uiMax};
              const r = await fetch('/api/manual_move', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
              if(!r.ok){ const t = await r.text(); showMessage('移動失敗: '+t,'error'); box2.classList.remove('shake'); void box2.offsetWidth; box2.classList.add('shake'); return; }
              const jd = await r.json(); if(jd && jd.result) scheduleCache[monthKey] = jd.result;
                  // record move for undo
                  pushChange({type:'move', data: {month: monthKey, resident: meta.resident, from_date: meta.from_date, from_hospital: meta.from_hospital, to_date: targetDate, to_hospital: it.hosp}});
                  showMessage('移動しました','success');
              renderCalendar(); if(document.getElementById('preview-area')) renderScheduleTable(scheduleCache[monthKey], 'preview-area'); if(document.getElementById('run-result')) renderScheduleTable(scheduleCache[monthKey], 'run-result');
              document.body.removeChild(overlay2);
            }catch(e){ showMessage('移動エラー: '+e.message,'error'); }
          });
          list2.appendChild(b);
        }
        box2.appendChild(list2);
        const cbtn = document.createElement('button'); cbtn.className='assign-cancel'; cbtn.textContent='キャンセル'; cbtn.addEventListener('click', ()=>{ document.body.removeChild(overlay2); });
        box2.appendChild(cbtn);
        overlay2.appendChild(box2);
        document.body.appendChild(overlay2);
      }catch(e){ showMessage('ドロップ処理エラー: '+e.message,'error'); }
    });
    row.appendChild(cell);
    if((first.getDay() + d) % 7 === 0){
      calGrid.appendChild(row);
      row = document.createElement('div'); row.className='cal-row';
    }
  }
  // append trailing row if any
  if(row.children.length) calGrid.appendChild(row);
}

// Bulk setters for calendar modal
function bulkSetForWeekdays(weekdayArray){
  const inputs = calGrid.querySelectorAll('.day-cell');
  for(const cell of inputs){
    const lbl = cell.querySelector('.cal-day-label').textContent;
    const d = parseInt(lbl,10);
    const dt = new Date(editingYear, editingMonth-1, d);
    if(weekdayArray.includes(dt.getDay())){
      const inp = cell.querySelector('.cal-day-input');
      const v = parseInt(document.getElementById('bulk-value').value || '0',10);
      inp.value = isNaN(v) ? 0 : v;
    }
  }
}

function bulkClearAll(){
  const inputs = calGrid.querySelectorAll('.day-cell');
  for(const cell of inputs){
    const inp = cell.querySelector('.cal-day-input');
    inp.value = 0;
  }
}

// attach bulk control handlers
document.getElementById('bulk-tue')?.addEventListener('click', ()=> bulkSetForWeekdays([2]));
document.getElementById('bulk-weekdays')?.addEventListener('click', ()=> bulkSetForWeekdays([1,2,3,4,5]));
document.getElementById('bulk-sat')?.addEventListener('click', ()=> bulkSetForWeekdays([6]));
document.getElementById('bulk-clear')?.addEventListener('click', ()=> bulkClearAll());

calPrev.addEventListener('click', ()=>{
  if(editingMonth===1){ editingMonth=12; editingYear--; } else { editingMonth--; }
  calTitle.textContent = `${editingHospital} — ${editingYear}/${String(editingMonth).padStart(2,'0')}`;
  fetchSchedule(`${editingYear}-${String(editingMonth).padStart(2,'0')}`).then(()=> renderCalendar());
});
calNext.addEventListener('click', ()=>{
  if(editingMonth===12){ editingMonth=1; editingYear++; } else { editingMonth++; }
  calTitle.textContent = `${editingHospital} — ${editingYear}/${String(editingMonth).padStart(2,'0')}`;
  fetchSchedule(`${editingYear}-${String(editingMonth).padStart(2,'0')}`).then(()=> renderCalendar());
});

calClose.addEventListener('click', ()=> closeCalendar());
calSave.addEventListener('click', ()=>{
  // read inputs and update currentCfg for editingHospital
  const inputs = calGrid.querySelectorAll('.day-cell');
  for(const cell of inputs){
    const lbl = cell.querySelector('.cal-day-label').textContent;
    const inp = cell.querySelector('.cal-day-input');
    const d = parseInt(lbl,10);
    const dateStr = `${editingYear}-${String(editingMonth).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const val = parseInt(inp.value||'0',10);
    if(!currentCfg[editingHospital]) currentCfg[editingHospital] = {};
    currentCfg[editingHospital][dateStr] = isNaN(val) ? 0 : val;
  }
  showMessage('カレンダーの変更を反映しました。');
  closeCalendar();
});

// Close modal when clicking outside inner area
calendarModal.addEventListener('click', (e)=>{
  if(e.target === calendarModal){
    closeCalendar();
  }
});

// Close modal on Escape key
document.addEventListener('keydown', (e)=>{
  if(e.key === 'Escape' && !calendarModal.classList.contains('hidden')){
    closeCalendar();
  }
});

document.getElementById('add-hospital').addEventListener('click', ()=> addHospitalRow());
document.getElementById('save-config').addEventListener('click', ()=> saveConfig());
document.getElementById('reload-config').addEventListener('click', ()=> loadConfig());
document.getElementById('run-solver').addEventListener('click', ()=> runSolver());
document.getElementById('download-xlsx')?.addEventListener('click', ()=> downloadXlsx());
document.getElementById('upload-both').addEventListener('click', ()=> uploadBoth());

// initial load
loadConfig();

async function uploadBoth(){
  const monthInput = document.getElementById('upload-month').value;
  if(!monthInput){ showMessage('対象月を選んでください','error'); return; }
  // monthInput is like YYYY-MM
  const f1 = document.getElementById('sheet1-file').files[0];
  const f2 = document.getElementById('sheet2-file').files[0];
  if(!f1 || !f2){ showMessage('両方のファイルを選択してください','error'); return; }
  const fd = new FormData();
  fd.append('month', monthInput);
  fd.append('sheet1', f1);
  fd.append('sheet2', f2);
  const resEl = document.getElementById('upload-result'); resEl.innerHTML = '解析中...';
  try{
    const res = await fetch('/api/upload_both', {method:'POST', body: fd});
    if(!res.ok){ const t = await res.text(); resEl.textContent = 'アップロード失敗: '+t; showMessage('アップロード失敗','error'); return; }
    const data = await res.json();
    // render results
    let html = '';
    if(data.status !== 'ok'){ resEl.textContent = '解析エラー'; return; }
    html += `<h4>解析結果</h4>`;
    html += `<div><strong>住民 (${data.residents.length})</strong><ul>`;
    for(const r of data.residents){ html += `<li>${r.name} (${r.rotation_type}) NG:${r.ng_dates.length}</li>`; }
    html += `</ul></div>`;
    html += `<div><strong>院外割当</strong><ul>`;
    for(const [d, arr] of Object.entries(data.assignments)){ html += `<li>${d}: ${arr.join(', ')}</li>`; }
    html += `</ul></div>`;
    if(data.unknown_names && data.unknown_names.length){ html += `<div><strong>未登録名</strong>: ${data.unknown_names.join(', ')}</div>`; }
    if(data.errors && (data.errors.sheet1.length || data.errors.sheet2.length)){ html += `<div><strong>パースエラー</strong>`; html += `<pre>${JSON.stringify(data.errors,null,2)}</pre></div>`; }
    resEl.innerHTML = html;
    showMessage('アップロード完了','success');
  }catch(e){ resEl.textContent = 'アップロード失敗: '+e.message; showMessage('エラー','error'); }
}

async function fetchSchedule(month){
  if(scheduleCache[month]) return scheduleCache[month];
  try{
    const res = await fetch(`/api/schedule?month=${encodeURIComponent(month)}`);
    if(!res.ok){
      const txt = await res.text();
      showMessage('スケジュール取得エラー: '+txt,'error');
      scheduleCache[month] = {status:'error'};
      return scheduleCache[month];
    }
    const data = await res.json();
    scheduleCache[month] = data;
    return data;
  }catch(e){
    showMessage('スケジュール取得失敗: '+e.message,'error');
    scheduleCache[month] = {status:'error'};
    return scheduleCache[month];
  }
}

async function fetchResidents(month){
  if(residentsCache[month]) return residentsCache[month];
  try{
    const res = await fetch(`/api/residents?month=${encodeURIComponent(month)}`);
    if(!res.ok) return null;
    const data = await res.json();
    residentsCache[month] = data.residents || [];
    return residentsCache[month];
  }catch(e){ return null; }
}

async function runSolver(){
  // determine month to run: prefer the upload-month input if present, else use current month
  const monthInput = document.getElementById('upload-month').value;
  let month = monthInput || null;
  if(!month){ const now = new Date(); month = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`; }
  const runBtn = document.getElementById('run-solver');
  const dlBtn = document.getElementById('download-xlsx');
  const runStatus = document.getElementById('run-status');
  // disable buttons and show spinner/status
  if(runBtn) runBtn.disabled = true;
  if(dlBtn) dlBtn.disabled = true;
  if(runStatus) runStatus.innerHTML = `<span class="spinner"></span>実行中...`;
  showMessage('スケジュール実行中... しばらくお待ちください');
  // safety timeout message if takes long
  let longWarnTimer = setTimeout(()=>{ if(runStatus) runStatus.textContent = '実行中（処理に時間がかかっています）'; }, 8000);
  try{
    const res = await fetch(`/api/run?month=${encodeURIComponent(month)}`, {method:'POST'});
    const txt = await res.text();
    if(!res.ok){
      showMessage('実行エラー: '+txt,'error');
      if(runStatus) runStatus.textContent = 'エラー';
      if(runBtn) runBtn.disabled = false;
      if(dlBtn) dlBtn.disabled = false;
      clearTimeout(longWarnTimer);
      return;
    }
    const data = JSON.parse(txt);
    if(data.status && data.status !== 'ok'){
      showMessage('実行結果: '+(data.message||'error'),'error');
      if(runStatus) runStatus.textContent = 'エラー';
      if(runBtn) runBtn.disabled = false;
      if(dlBtn) dlBtn.disabled = false;
      clearTimeout(longWarnTimer);
      return;
    }
    const out = data.result || data;
    renderPreview(out);
    // also render the table under the Run Solver button
    renderScheduleTable(out, 'run-result');
    showMessage('スケジュールを生成・保存しました。','success');
    if(runStatus) runStatus.textContent = '完了';
    if(runBtn) runBtn.disabled = false;
    if(dlBtn) dlBtn.disabled = false;
    clearTimeout(longWarnTimer);
  }catch(e){
    showMessage('実行失敗: '+e.message,'error');
    const runStatus = document.getElementById('run-status');
    if(runStatus) runStatus.textContent = 'エラー';
    const runBtn = document.getElementById('run-solver');
    const dlBtn = document.getElementById('download-xlsx');
    if(runBtn) runBtn.disabled = false;
    if(dlBtn) dlBtn.disabled = false;
    clearTimeout(longWarnTimer);
  }
}

function downloadXlsx(){
  const monthInput = document.getElementById('upload-month').value;
  let month = monthInput || null;
  if(!month){ const now = new Date(); month = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`; }
  const url = `/api/download?month=${encodeURIComponent(month)}`;
  // open download in new tab/window to trigger attachment
  window.open(url, '_blank');
}

function renderPreview(out){
  // render into calendar preview area
  renderScheduleTable(out, 'preview-area');
}

// Render schedule table into a container by id
function renderScheduleTable(out, containerId){
  const container = document.getElementById(containerId);
  if(!container) return;
  container.innerHTML = '';
  if(!out || !out.dates || !out.hospitals){
    container.textContent = 'スケジュールデータがありません。';
    return;
  }
  // determine monthKey for resident lookups
  let monthKey = out.month;
  if(!monthKey && out.dates && out.dates.length){ try{ monthKey = out.dates[0].slice(0,7); }catch(e){ monthKey = null; } }
  if(!monthKey){ const uiMonth = document.getElementById('upload-month')?.value; if(uiMonth) monthKey = uiMonth; else { const now = new Date(); monthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`; } }

  const table = document.createElement('table');
  table.className = 'preview-table';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const thDate = document.createElement('th'); thDate.textContent = '';
  headRow.appendChild(thDate);
  for(const h of out.hospitals){ const th = document.createElement('th'); th.textContent = h; headRow.appendChild(th); }
  thead.appendChild(headRow);
  table.appendChild(thead);
  const tbody = document.createElement('tbody');
  for(const d of out.dates){
    const tr = document.createElement('tr');
    const dt = new Date(d);
    const tdDate = document.createElement('td'); tdDate.textContent = (dt.getMonth()+1) + '月' + dt.getDate() + '日'; tr.appendChild(tdDate);
    for(const h of out.hospitals){
      const td = document.createElement('td');
      td.dataset.date = d; td.dataset.hospital = h;
      const arr = (out.assignments && out.assignments[d] && out.assignments[d][h]) || [];
      // for each assigned name, create draggable label
      const cellBox = document.createElement('div'); cellBox.className = 'preview-cell-box';
      for(const name of arr){
        const span = document.createElement('div'); span.className = 'assigned-label'; span.textContent = name; span.draggable = true;
        // drag handlers
        span.addEventListener('dragstart', (ev)=>{
          const meta = {resident: name, from_date: d, from_hospital: h};
          currentDragMeta = meta;
          try{ ev.dataTransfer.setData('application/json', JSON.stringify(meta)); ev.dataTransfer.effectAllowed = 'move'; }catch(e){}
        });
        span.addEventListener('dragend', (ev)=>{ currentDragMeta = null; });
        // click selects resident and highlights preview targets
        span.addEventListener('click', async ()=>{
          // fetch resident ng_dates
          const rs = await fetchResidents(monthKey);
          const found = rs ? rs.find(r => r.name === name) : null;
          residentHighlight = {name: name, month: monthKey, ng_dates: found ? (found.ng_dates || []) : []};
          highlightPreviewTargets(monthKey, residentHighlight.name);
          showMessage(`${name} を選択しました。移動可能箇所がハイライトされます。`,'info');
        });
        cellBox.appendChild(span);
      }
      td.appendChild(cellBox);
      // attach drop handlers to td
      td.addEventListener('dragover', (e)=>{
        if(!currentDragMeta) return;
        // prevent drop on NG dates for this resident
        const ngDates = residentHighlight && residentHighlight.name === currentDragMeta.resident ? residentHighlight.ng_dates || [] : [];
        const isNg = Array.isArray(ngDates) && ngDates.indexOf(d) !== -1;
        if(isNg){ td.classList.add('drop-not-allowed'); return; }
        e.preventDefault(); td.classList.add('drop-target');
      });
      td.addEventListener('dragleave', (e)=>{ td.classList.remove('drop-target'); td.classList.remove('drop-not-allowed'); });
      td.addEventListener('drop', async (e)=>{
        if(!currentDragMeta) return;
        e.preventDefault(); td.classList.remove('drop-target'); td.classList.remove('drop-not-allowed');
        const meta = currentDragMeta;
        // prevent dropping onto same date/hospital
        if(meta.from_date === d && meta.from_hospital === h){ showMessage('同じ場所です','error'); return; }
        // ensure not NG for resident
        const rs = await fetchResidents(monthKey);
        const found = rs ? rs.find(r=>r.name === meta.resident) : null;
        const isNg = found && Array.isArray(found.ng_dates) && found.ng_dates.indexOf(d) !== -1;
        if(isNg){ showMessage('NG日にはドロップできません','error'); return; }
        // before calling manual_move, check capacity/remaining for this target cell
        try{
          // compute configured capacity (date-specific > weekday key), default 0
          let capacity = 0;
          try{
            const cfg = currentCfg[h] || {};
            if(cfg[d] !== undefined){ capacity = Number(cfg[d]) || 0; }
            else {
              const dt2 = new Date(d);
              const jsDay2 = dt2.getDay();
              const wk2 = (jsDay2 + 6) % 7;
              if(cfg[String(wk2)] !== undefined){ capacity = Number(cfg[String(wk2)]) || 0; }
            }
          }catch(e){ capacity = 0; }
          const assignedArr2 = (sched.assignments && sched.assignments[d] && sched.assignments[d][h]) || [];
          const assignedCount2 = Array.isArray(assignedArr2) ? assignedArr2.length : 0;
          // check holiday for university extra slots
          let isHoliday2 = false;
          try{ const hres2 = await fetch(`/api/is_holiday?date=${encodeURIComponent(d)}`); if(hres2.ok){ const hj2 = await hres2.json(); isHoliday2 = !!hj2.is_holiday; } }catch(e){}
          const dt3 = new Date(d); const isWeekend3 = dt3.getDay() === 0 || dt3.getDay() === 6;
          if(h === '大学病院' && (isWeekend3 || isHoliday2)) capacity = capacity + 4;
          const remaining2 = capacity - assignedCount2;
          if(remaining2 <= 0 && Array.isArray(assignedArr2) && assignedArr2.indexOf(meta.resident) === -1){
            showMessage('この病院にはスロットがありません','error');
            return;
          }

          const uiMax = Number(document.getElementById('max-assignments')?.value || 0) || 2;
          const payload = {month: monthKey, resident: meta.resident, from_date: meta.from_date, from_hospital: meta.from_hospital, to_date: d, to_hospital: h, max_assignments: uiMax};
          const res = await fetch('/api.manual_move' in window ? '/api.manual_move' : '/api/manual_move', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
          if(!res.ok){ const t = await res.text(); showMessage('移動失敗: '+t,'error'); return; }
          const jd = await res.json(); if(jd && jd.result) scheduleCache[monthKey] = jd.result;
          pushChange({type:'move', data: {month: monthKey, resident: meta.resident, from_date: meta.from_date, from_hospital: meta.from_hospital, to_date: d, to_hospital: h}});
          showMessage('移動しました','success');
          renderScheduleTable(scheduleCache[monthKey], containerId);
        }catch(e){ showMessage('移動処理エラー: '+e.message,'error'); }
      });
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  container.appendChild(table);
  // If per-res assignment counts are present, render a summary below the table
  if(out.per_res_counts && out.per_res_required){
    const sumDiv = document.createElement('div');
    sumDiv.className = 'per-res-summary';
    const h = document.createElement('h4'); h.textContent = '各人の割当数（実績 / 目標）'; sumDiv.appendChild(h);
    const t = document.createElement('table'); t.className = 'summary-table';
    const th = document.createElement('thead');
    const thr = document.createElement('tr');
    ['名前','実績','目標'].forEach(x=>{ const c=document.createElement('th'); c.textContent=x; thr.appendChild(c); });
    th.appendChild(thr); t.appendChild(th);
    const tb = document.createElement('tbody');
    const names = Object.keys(out.per_res_required).sort();
    // determine month string for resident lookup
    let monthKey = out.month;
    if(!monthKey && out.dates && out.dates.length) {
      // derive month from first date
      try{ monthKey = out.dates[0].slice(0,7); }catch(e){ monthKey = null; }
    }
    // fallback: use selected upload-month input or current month
    if(!monthKey){
      const uiMonth = document.getElementById('upload-month')?.value;
      if(uiMonth){ monthKey = uiMonth; }
      else { const now = new Date(); monthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`; }
    }
    for(const name of names){
      const tr = document.createElement('tr');
      const tdName = document.createElement('td'); tdName.textContent = name; tr.appendChild(tdName);
      const actual = (out.per_res_counts[name] !== undefined) ? Number(out.per_res_counts[name]) : 0;
      const target = Number(out.per_res_required[name] || 0);
      const tdAct = document.createElement('td'); tdAct.textContent = String(actual); tr.appendChild(tdAct);
      const tdReq = document.createElement('td'); tdReq.textContent = String(target); tr.appendChild(tdReq);
      // highlight when actual < target
      if(target > 0 && actual < target){
        tr.classList.add('not-met');
      }
      // click name to show resident NG calendar
      tdName.classList.add('res-name');
      tdName.style.cursor = 'pointer';
      tdName.title = 'クリックしてNG日を表示';
      tdName.addEventListener('click', ()=> showResidentCalendar(name, monthKey));
      tb.appendChild(tr);
    }
    t.appendChild(tb);
    sumDiv.appendChild(t);
    container.appendChild(sumDiv);
  }
}

function highlightPreviewTargets(monthKey, residentName){
  // ensure residents data present
  (async ()=>{
    const rs = await fetchResidents(monthKey);
    const found = rs ? rs.find(r=>r.name === residentName) : null;
    const ngDates = found ? (found.ng_dates || []) : [];
    const container = document.querySelector('.preview-table');
    if(!container) return;
    // clear previous highlights
    const tds = container.querySelectorAll('td');
    tds.forEach(td=> td.classList.remove('preview-drop-target','ng-preview'));
    // iterate td and mark
    for(const td of container.querySelectorAll('td')){
      const d = td.dataset.date;
      const h = td.dataset.hospital;
      if(!d || !h) continue;
      if(Array.isArray(ngDates) && ngDates.indexOf(d) !== -1){ td.classList.add('ng-preview'); continue; }
      // compute capacity remaining for this hospital/date using currentCfg and scheduleCache
      const sched = scheduleCache[monthKey] || {};
      let capacity = 0;
      try{
        const cfg = currentCfg[h] || {};
        if(cfg[d] !== undefined){ capacity = Number(cfg[d]) || 0; }
        else { const dt = new Date(d); const jsDay = dt.getDay(); const wk = (jsDay + 6) % 7; if(cfg[String(wk)] !== undefined) capacity = Number(cfg[String(wk)]) || 0; }
      }catch(e){ capacity = 0; }
      const assignedArr = (sched.assignments && sched.assignments[d] && sched.assignments[d][h]) || [];
      const assignedCount = Array.isArray(assignedArr) ? assignedArr.length : 0;
      const dt = new Date(d); const isWeekend = dt.getDay() === 0 || dt.getDay() === 6;
      if(h === '大学病院' && (isWeekend || false)) capacity = capacity + 4;
      const remaining = capacity - assignedCount;
      if(remaining > 0) td.classList.add('preview-drop-target');
    }
  })();
}

async function showResidentCalendar(name, month){
  if(!month){ showMessage('月情報がありません','error'); return; }
  try{
    const res = await fetch(`/api/residents?month=${encodeURIComponent(month)}`);
    if(!res.ok){ throw new Error(await res.text()); }
    const data = await res.json();
    const residents = data.residents || [];
    const found = residents.find(r => r.name === name);
    if(!found){ showMessage('該当住民データが見つかりません','error'); return; }
    residentHighlight = {name: name, month: month, ng_dates: found.ng_dates || []};
    // show calendar modal focused on month
    const parts = month.split('-');
    editingYear = Number(parts[0]); editingMonth = Number(parts[1]);
    editingHospital = null; // not editing hospital slots
    calTitle.textContent = `${name} — NG日 (${month})`;
    await fetchSchedule(month);
    renderCalendar();
    calendarModal.classList.remove('hidden');
    // ensure undo button exists above calendar grid
    ensureUndoButton();
    // hide save controls when viewing resident NGs
    // (we keep inputs but user should not save changes in this view)
  }catch(e){ showMessage('住民取得失敗: '+e.message,'error'); }
}

// assignment modal: choose hospital and post manual assignment
async function openAssignModal(dateIso, residentName, month){
  // build modal DOM
  const overlay = document.createElement('div'); overlay.className = 'assign-modal-overlay';
  const box = document.createElement('div'); box.className = 'assign-modal-box';
  const h = document.createElement('h3'); h.textContent = `${residentName} を ${dateIso} に割当`;
  box.appendChild(h);
  const info = document.createElement('div'); info.textContent = '病院を選択してください'; box.appendChild(info);
  // inline error area inside modal
  const err = document.createElement('div'); err.className = 'assign-error'; err.style.display = 'none'; box.appendChild(err);

  // ensure up-to-date schedule data for that month
  if(!scheduleCache[month]){
    await fetchSchedule(month);
  }
  const sched = scheduleCache[month] || {};
  // determine candidate hospitals: prefer scheduleCache month hospitals, else currentCfg keys
  let hospitals = [];
  if(sched && sched.hospitals && Array.isArray(sched.hospitals)) hospitals = sched.hospitals.slice();
  if(!hospitals || hospitals.length===0) hospitals = Object.keys(currentCfg || {});
  if(!hospitals || hospitals.length===0){ showMessage('病院リストが取得できません','error'); return; }

  // filter hospitals by capacity for the given date (compute capacity from currentCfg)
  // determine holiday status for date (ask server)
  let isHoliday = false;
  try{
    const hres = await fetch(`/api/is_holiday?date=${encodeURIComponent(dateIso)}`);
    if(hres.ok){ const hj = await hres.json(); isHoliday = !!hj.is_holiday; }
  }catch(e){ isHoliday = false; }

  const candidates = [];
  for(const hosp of hospitals){
    if(!hosp) continue; // skip falsy/null names
    // compute configured capacity for date
    let capacity = 0;
    try{
      const cfg = currentCfg[hosp] || {};
      if(cfg[dateIso] !== undefined){ capacity = Number(cfg[dateIso]) || 0; }
      else {
        const dt = new Date(dateIso);
        const jsDay = dt.getDay(); // 0=Sun..6=Sat
        const wk = (jsDay + 6) % 7; // map to 0=Mon..6=Sun like backend
        if(cfg[String(wk)] !== undefined){ capacity = Number(cfg[String(wk)]) || 0; }
      }
    }catch(e){ capacity = 0; }
    // assigned count on that date
    const assignedArr = (sched.assignments && sched.assignments[dateIso] && sched.assignments[dateIso][hosp]) || [];
    const assignedCount = Array.isArray(assignedArr) ? assignedArr.length : 0;
    // If university hospital on weekend or holiday, increase capacity by 4
    const dt = new Date(dateIso);
    const isWeekend = dt.getDay() === 0 || dt.getDay() === 6;
    if(hosp === '大学病院' && (isWeekend || isHoliday)){
      capacity = capacity + 4;
    }
    const remaining = capacity - assignedCount;
    if(remaining > 0){ candidates.push({hosp, remaining}); }
  }

  // determine if resident is already assigned on this date (collect hospitals)
  const assignedHospitals = [];
  if(sched && sched.assignments && sched.assignments[dateIso]){
    for(const hh of Object.keys(sched.assignments[dateIso] || {})){
      const arr = sched.assignments[dateIso][hh] || [];
      if(Array.isArray(arr) && arr.indexOf(residentName) !== -1){ assignedHospitals.push(hh); }
    }
  }

  // if there are no candidates and resident is not assigned, show message
  if(candidates.length === 0 && assignedHospitals.length === 0){
    showMessage('該当日の残スロットがありません','error');
    return;
  }

    const list = document.createElement('div'); list.className = 'assign-list';
    // If resident already has assignment(s) on this date, show a delete button
    if(assignedHospitals.length){
      const del = document.createElement('button'); del.className = 'assign-delete'; del.textContent = '予定を削除する';
      del.addEventListener('click', async ()=>{
        try{
          const payload = {month: month, date: dateIso, resident: residentName};
          const res = await fetch('/api/manual_unassign', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
          if(!res.ok){ const txt = await res.text(); err.textContent = '削除失敗: '+txt; err.style.display = 'block'; box.classList.remove('shake'); void box.offsetWidth; box.classList.add('shake'); return; }
          const data = await res.json();
          if(data && data.result){ scheduleCache[month] = data.result; }
          // record unassign for undo (record hospitals removed)
          pushChange({type:'unassign', data: {month: month, date: dateIso, resident: residentName, hospitals: assignedHospitals.slice()}});
          showMessage('予定を削除しました','success');
          // refresh UI
          renderCalendar();
          if(document.getElementById('preview-area')) renderScheduleTable(scheduleCache[month], 'preview-area');
          if(document.getElementById('run-result')) renderScheduleTable(scheduleCache[month], 'run-result');
          document.body.removeChild(overlay);
        }catch(e){ err.textContent = '削除エラー: '+e.message; err.style.display = 'block'; box.classList.remove('shake'); void box.offsetWidth; box.classList.add('shake'); }
      });
      list.appendChild(del);
    }
  for(const item of candidates){
    const hosp = item.hosp;
    const remaining = item.remaining;
    const btn = document.createElement('button'); btn.className = 'assign-btn'; btn.innerHTML = `${hosp}<br><small>残${remaining}</small>`;
    btn.addEventListener('click', async ()=>{
      try{
        // read max assignments from UI (fallback 2)
        const uiMax = Number(document.getElementById('max-assignments')?.value || 0) || 2;
        // compute current count from cache (fallback to 0)
        const curCounts = (scheduleCache[month] && scheduleCache[month].per_res_counts) || {};
        const cur = Number(curCounts[residentName] || 0);
        // check current assignments on this date: if resident already assigned that day, replacing won't increase count
        let assigned_on_date = false;
        if(scheduleCache[month] && scheduleCache[month].assignments && scheduleCache[month].assignments[dateIso]){
          for(const hh of Object.keys(scheduleCache[month].assignments[dateIso] || {})){
            const arr = scheduleCache[month].assignments[dateIso][hh] || [];
            if(Array.isArray(arr) && arr.indexOf(residentName) !== -1){ assigned_on_date = true; break; }
          }
        }
        if((cur - (assigned_on_date ? 1 : 0) + 1) > uiMax){
          // show inline error and shake modal
          err.textContent = `上限回数（${uiMax}回）に達しています`;
          err.style.display = 'block';
          box.classList.remove('shake');
          // trigger reflow to re-run animation
          void box.offsetWidth;
          box.classList.add('shake');
          return;
        }

        const payload = {month: month, date: dateIso, resident: residentName, hospital: hosp, max_assignments: uiMax};
        const res = await fetch('/api/manual_assign', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
        if(!res.ok){ const txt = await res.text();
          err.textContent = '割当失敗: '+txt;
          err.style.display = 'block';
          box.classList.remove('shake'); void box.offsetWidth; box.classList.add('shake');
          return; }
        const data = await res.json();
        // update cache with returned solver result to reflect assignment immediately
        if(data && data.result){ scheduleCache[month] = data.result; }
        // record change for undo
        pushChange({type:'assign', data: {month: month, date: dateIso, resident: residentName, hospital: hosp}});
        showMessage('割当を保存しました','success');
        // refresh schedule and calendar preview
        if(!scheduleCache[month]) await fetchSchedule(month);
        renderCalendar();
        if(document.getElementById('preview-area')){
          const out = scheduleCache[month];
          renderScheduleTable(out, 'preview-area');
        }
        if(document.getElementById('run-result')){
          const out = scheduleCache[month];
          renderScheduleTable(out, 'run-result');
        }
      }catch(e){ showMessage('割当エラー: '+e.message,'error'); }
      // close modal
      document.body.removeChild(overlay);
    });
    list.appendChild(btn);
  }
  box.appendChild(list);
  const cancel = document.createElement('button'); cancel.className='assign-cancel'; cancel.textContent='キャンセル'; cancel.addEventListener('click', ()=>{ document.body.removeChild(overlay); });
  box.appendChild(cancel);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
}
