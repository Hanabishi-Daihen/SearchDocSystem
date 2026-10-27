/**
 * Sync: เทียบรายการไฟล์ใน OneDrive กับ sheet Files → แตกข้อความไฟล์ใหม่/ที่เปลี่ยน → สร้าง search index
 *   runScheduled(e)  เรียกโดย time trigger
 *   syncNow()        เรียกจากเมนูใน Google Sheet
 * งานที่ทำไม่ทันในรอบเดียว (GAS รันได้ครั้งละ 6 นาที) จะทำต่อในรอบถัดไป
 */

function runScheduled(e) {
  const isTrigger = e && e.triggerUid &&
    ScriptApp.getProjectTriggers().some((t) => t.getUniqueId() === e.triggerUid);
  if (!isTrigger) requireOwner_();
  return runSync_();
}

function syncNow() {
  requireOwner_();
  return runSync_();
}

function runSync_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    log_('INFO', 'ข้ามรอบนี้: มีงาน sync อื่นกำลังทำงานอยู่');
    flushLogs_();
    return { skipped: true };
  }
  const started = Date.now();
  const summary = { total: 0, added: 0, changed: 0, moved: 0, deleted: 0, extracted: 0, failed: 0, remaining: 0 };
  try {
    const files = new Table(getSpreadsheet_(), 'Files');
    const diff = applyListing_(files, listAllFiles_(getSource_()));
    Object.assign(summary, diff);
    files.save();

    const processed = processQueue_(files, started);
    summary.extracted = processed.extracted;
    summary.failed = processed.failed;
    files.save();

    summary.remaining = files.filter((f) => f.status === STATUS.QUEUED).length;
    summary.total = files.filter((f) => f.status !== STATUS.DELETED).length;
    if (diff.added || diff.changed || diff.moved || diff.deleted || processed.extracted || processed.failed ||
        !props_().getProperty('INDEX_MANIFEST')) {
      rebuildSearchIndex_(files);
    }
    props_().setProperty('LAST_SYNC', JSON.stringify({ at: nowIso_(), ok: true, total: summary.total, remaining: summary.remaining }));
    log_('INFO', `sync เสร็จใน ${Math.round((Date.now() - started) / 1000)} วินาที: ${JSON.stringify(summary)}`);
    return summary;
  } catch (err) {
    props_().setProperty('LAST_SYNC', JSON.stringify({ at: nowIso_(), ok: false, error: err.message }));
    log_('ERROR', `sync ล้มเหลว: ${err.message}`);
    throw err;
  } finally {
    flushLogs_();
    lock.releaseLock();
  }
}

/** เทียบรายการจาก OneDrive กับตาราง Files (แก้ใน memory) */
function applyListing_(files, listing) {
  const settings = getSettings_();
  const result = { added: 0, changed: 0, moved: 0, deleted: 0 };
  const seen = new Set();

  listing.forEach(({ item, path }) => {
    seen.add(item.id);
    const ext = (item.name.match(/\.([A-Za-z0-9]{1,6})$/) || ['', ''])[1].toLowerCase();
    const changeKey = item.cTag || item.eTag || `${item.fileSystemInfo && item.fileSystemInfo.lastModifiedDateTime}|${item.size}`;
    const extractable = settings.fileExtensions.includes(ext) && Number(item.size) <= settings.maxFileBytes;
    const meta = {
      name: item.name,
      ext,
      path,
      web_url: item.webUrl || '',
      size: item.size || 0,
      modified_at: (item.fileSystemInfo && item.fileSystemInfo.lastModifiedDateTime) || '',
      modified_by: (item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName) || '',
      change_key: changeKey,
    };

    const row = files.get(item.id);
    if (!row) {
      files.insert(Object.assign({ id: item.id, first_seen_at: nowIso_(), attempts: 0 }, meta,
        { status: extractable ? STATUS.QUEUED : STATUS.NAME_ONLY }));
      result.added++;
      return;
    }
    if (row.name !== meta.name || row.path !== meta.path || row.web_url !== meta.web_url) result.moved++;
    const contentChanged = row.indexed_key !== changeKey;
    const wasDeleted = row.status === STATUS.DELETED;
    files.update(row, meta);
    if (!extractable) {
      if (row.status !== STATUS.NAME_ONLY) files.update(row, { status: STATUS.NAME_ONLY });
    } else if (contentChanged && row.status !== STATUS.QUEUED) {
      // เนื้อหาเปลี่ยน / ไฟล์กลับมา / เพิ่งรองรับชนิดไฟล์นี้ → แตกข้อความใหม่
      // (ไฟล์ที่ล้มเหลวครบแล้วมี indexed_key = key ตอนล้มเหลว จึงไม่วนลองใหม่จนกว่าไฟล์จะถูกแก้)
      files.update(row, { status: STATUS.QUEUED, attempts: 0, last_error: '' });
      result.changed++;
    } else if (!contentChanged && wasDeleted) {
      files.update(row, { status: row.text_file_id ? STATUS.INDEXED : STATUS.QUEUED });
    }
  });

  files.rows.forEach((row) => {
    if (!seen.has(row.id) && row.status !== STATUS.DELETED) {
      files.update(row, { status: STATUS.DELETED });
      result.deleted++;
    }
  });
  return result;
}

/** แตกข้อความไฟล์ในคิวจนกว่าจะใกล้หมดเวลา */
function processQueue_(files, started) {
  const source = getSource_();
  const out = { extracted: 0, failed: 0 };
  const queue = files.filter((f) => f.status === STATUS.QUEUED)
    .sort((a, b) => b.modified_at.localeCompare(a.modified_at)); // ไฟล์ใหม่ก่อน
  for (const row of queue) {
    if (Date.now() - started > MAX_RUNTIME_MS) break;
    const key = row.change_key;
    try {
      const blob = downloadItem_(source.driveId, row.id);
      const text = extractText_(blob, row.ext).replace(/\r\n?/g, '\n').trim();
      files.update(row, {
        text_file_id: saveText_(row, text),
        text_chars: text.length,
        status: STATUS.INDEXED,
        indexed_key: key,
        indexed_at: nowIso_(),
        attempts: 0,
        last_error: '',
      });
      out.extracted++;
    } catch (err) {
      const attempts = Number(row.attempts || 0) + 1;
      files.update(row, {
        attempts,
        last_error: String(err.message).slice(0, 500),
        status: attempts >= MAX_ATTEMPTS ? STATUS.ERROR : STATUS.QUEUED,
        // ถ้าล้มเหลวครบแล้ว จำ key ไว้ จะไม่ลองใหม่จนกว่าไฟล์จะถูกแก้
        indexed_key: attempts >= MAX_ATTEMPTS ? key : row.indexed_key,
      });
      log_('WARN', `แตกข้อความ "${row.name}" ไม่สำเร็จ (ครั้งที่ ${attempts}): ${err.message}`);
      out.failed++;
    }
  }
  return out;
}

/** เมนู: ให้ไฟล์ที่ล้มเหลวลองใหม่ */
function retryFailed() {
  requireOwner_();
  const files = new Table(getSpreadsheet_(), 'Files');
  let n = 0;
  files.filter((f) => f.status === STATUS.ERROR).forEach((f) => {
    files.update(f, { status: STATUS.QUEUED, attempts: 0, indexed_key: '' });
    n++;
  });
  files.save();
  return n;
}
