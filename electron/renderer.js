const dirInput = document.getElementById('dir');
const browseBtn = document.getElementById('browse');
const startBtn = document.getElementById('start');
const recursiveChk = document.getElementById('recursive');
const overwriteChk = document.getElementById('overwrite');
const dryRunChk = document.getElementById('dryRun');
const logEl = document.getElementById('log');
const summaryEl = document.getElementById('summary');

function appendLog(line) {
  const div = document.createElement('div');
  div.textContent = line;
  logEl.appendChild(div);
  logEl.scrollTop = logEl.scrollHeight;
}

function tag(t) {
  const span = document.createElement('span');
  span.className = 'tag ' + (t === 'ok' ? 'ok' : t === 'fail' ? 'fail' : t === 'skip' ? 'skip' : 'dry');
  span.textContent = t;
  return span;
}

browseBtn.addEventListener('click', async () => {
  const dir = await window.api.pickDirectory();
  if (dir) dirInput.value = dir;
});

let unsubscribe = null;
startBtn.addEventListener('click', async () => {
  startBtn.disabled = true;
  logEl.innerHTML = '';
  summaryEl.textContent = '';

  if (unsubscribe) { unsubscribe(); unsubscribe = null; }
  unsubscribe = window.api.onProgress((evt) => {
    const line = document.createElement('div');
    line.appendChild(tag(evt.type));
    line.appendChild(document.createTextNode(' ' + evt.file + ' => ' + evt.pdf + (evt.error ? (' [' + evt.error + ']') : '')));
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  });

  const payload = {
    dir: dirInput.value,
    recursive: !!recursiveChk.checked,
    overwrite: !!overwriteChk.checked,
    dryRun: !!dryRunChk.checked
  };
  const res = await window.api.runOfficeToPdfMissing(payload);
  startBtn.disabled = false;
  if (!res || !res.ok) {
    appendLog('[错误] ' + (res && res.error ? res.error : '未知错误'));
    return;
  }
  const s = res.summary;
  summaryEl.textContent = `总计 ${s.total}，待处理 ${s.candidates}，成功 ${s.converted}${s.dryRun ? ' (dry-run)' : ''}，跳过 ${s.skippedExist}，失败 ${s.failed}`;
});

