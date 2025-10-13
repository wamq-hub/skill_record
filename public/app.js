// ============================================
// السكريبت الرئيسي - app.js (نسخة مصححة)
// ============================================

let currentUser = null;

// ✅ اجعل هذه المتغيرات أعلى الملف (لتخزين آخر بيانات جلبتها لوحاتك):
window.trainerWorkshops = [];
window.pendingAttendanceCache = [];
window.traineeSkillsCache = []; // سيُملأ في لوحة المتدرب

// ---- أدوات مساعدة واجهة ----
function showError(message) {
  const errorMsg = document.getElementById('errorMessage');
  if (!errorMsg) return alert(message);
  errorMsg.textContent = message;
  errorMsg.classList.add('show');
}
function normalizeUserType(v){ return (v||'').replace(/\s+/g,'_'); }

if (typeof ExcelJS === 'undefined') {
  console.error('❌ ExcelJS غير محمّل! أضف السكريبت التالي في index.html قبل app.js:\n<script src="https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js"></script>');
  alert('⚠️ مكتبة ExcelJS غير محملة!\n\nأضف السكريبت التالي في index.html:\n<script src="https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js"></script>');
}


// ===== أوفرلاي تحميل الملف =====
let __fileOverlayStart = 0;
function showFileOverlay(msg = 'جاري تجهيز الملف...') {
  const ov = document.getElementById('fileDownloadOverlay');
  const m  = document.getElementById('fileOverlayMsg');
  if (!ov || !m) { 
    console.error('❌ fileDownloadOverlay غير موجود في DOM'); // سطر خطأ
    return;
  }
  __fileOverlayStart = Date.now();
  m.textContent = msg;
  ov.classList.remove('hidden');
}
function updateFileOverlay(msg) {
  const m = document.getElementById('fileOverlayMsg');
  if (m) m.textContent = msg;
}
async function hideFileOverlay(minMs = 1200) {
  const ov = document.getElementById('fileDownloadOverlay');
  const diff = Date.now() - __fileOverlayStart;
  const waitMore = Math.max(0, minMs - diff);
  await new Promise(r => setTimeout(r, waitMore));
  ov?.classList.add('hidden');
}


// ---- تسجيل الدخول ----
async function handleLogin(event) {
  event.preventDefault();

  const usernameEl = document.getElementById('username');
  const spinner = document.getElementById('loadingSpinner');
  const username = (usernameEl?.value || '').trim();

  if (!username) return showError('الرجاء إدخال رقم الهوية أو اسم المستخدم');

  spinner?.classList.remove('hidden');
  document.getElementById('errorMessage')?.classList.remove('show');

  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=login&username=${encodeURIComponent(username)}`;
    const res = await fetch(url);
    const data = await res.json();

    if (data?.success) {
      // ✅ التصحيح هنا
      currentUser = { ...data.user, userType: normalizeUserType(data.user.userType) };
      localStorage.setItem('currentUser', JSON.stringify(currentUser));
      redirectToDashboard(currentUser.userType);
    } else {
      showError(data?.message || 'اسم المستخدم غير موجود');
    }
  } catch (err) {
    console.error('خطأ في تسجيل الدخول:', err);
    showError('حدث خطأ في الاتصال. الرجاء المحاولة مرة أخرى');
  } finally {
    spinner?.classList.add('hidden');
  }
}


// ---- توجيه للوحة المناسبة (نسخة واحدة فقط) ----
function redirectToDashboard(userTypeRaw) {
  const t = normalizeUserType(userTypeRaw);
  document.getElementById('loginPage')?.classList.add('hidden');

  switch (t) {
    case CONFIG.USER_TYPES.TRAINEE:  return showTraineeDashboard();
    case CONFIG.USER_TYPES.TRAINER:  return showTrainerDashboard();
    case CONFIG.USER_TYPES.HEAD:     return showHeadDashboard();
    case CONFIG.USER_TYPES.DEAN_STUDENTS:
    case CONFIG.USER_TYPES.DEAN:     return showAdminDashboard();
    default:
      // فاصل أمان: لا تترك الصفحة بيضاء
      showError('⚠️ نوع الحساب غير معروف. تأكد من قيمة نوع المستخدم في الشيت.');
      document.getElementById('loginPage')?.classList.remove('hidden');
      localStorage.removeItem('currentUser');
  }
}

// ---- لوحة المتدرب ----
async function showTraineeDashboard() {
  document.getElementById('traineePage')?.classList.remove('hidden');
  document.getElementById('traineeInfo').textContent = `${currentUser.name} - ${currentUser.studentId}`;
  await loadTraineeData();
}

async function loadTraineeData() {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getTraineeData&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    const data = await res.json();
    if (data?.success) {
      window.traineeSkillsCache = data.skills || []; // ✅ مهم للشهادات
      renderTraineeStats(data.stats);
      renderAvailableWorkshops(data.workshops);
      renderTraineeSkills(data.skills);
      await loadTraineeTrack(); // ← يجلب المسار ويعرضه
    } else {
      showError(data?.message || 'تعذّر جلب بيانات المتدرب');
    }
  } catch (e) {
    console.error('خطأ في جلب البيانات:', e);
    showError('حدث خطأ في جلب البيانات');
  } finally {
    spinner?.classList.add('hidden');
  }
}

async function loadTraineeTrack() {
  const holder = document.getElementById('traineeTrack');
  if (!holder) {
    console.error('❌ عنصر traineeTrack غير موجود في الصفحة'); // سطر خطأ
    return;
  }
  holder.innerHTML = '<p style="color:var(--tvtc-text-muted)">جاري التحميل...</p>';

  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getTraineeTrack&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    const data = await res.json();

    if (data?.success && Array.isArray(data.track)) {
      renderTraineeTrack(data.track);
    } else if (data && data.message === 'Invalid action') {
      // سطر تصحيح: بديل آمن يعتمد على المهارات المعتمدة فقط
      console.warn('ℹ️ لا يوجد getTraineeTrack في GAS — سيتم إنشاء مسار مبسّط من المهارات المعتمدة.');
      const approved = (window.traineeSkillsCache||[])
        .filter(s => s.status === 'معتمد')
        .map(s => ({ workshopName:s.name, date:s.date, status:'معتمد', registrationTime:'—' }));
      renderTraineeTrack(approved, /*isFallback=*/true);
    } else {
      console.error('❌ استجابة غير متوقعة لمسار التتبّع:', data); // سطر خطأ
      holder.innerHTML = '<p style="color:#b91c1c">تعذّر جلب مسار التتبّع.</p>';
    }
  } catch (e) {
    console.error('❌ خطأ شبكة أثناء جلب التتبّع:', e); // سطر خطأ
    document.getElementById('traineeTrack').innerHTML =
      '<p style="color:#b91c1c">تعذّر الاتصال لجلب مسار التتبّع.</p>';
  }
}

function renderTraineeTrack(items, isFallback=false) {
  const el = document.getElementById('traineeTrack');
  if (!items || !items.length) {
    el.innerHTML = '<p style="color:var(--tvtc-text-muted)">لا توجد ورش مسجّل بها حتى الآن.</p>';
    return;
  }

  const badge = (status) => {
    if (status === 'معتمد') return 'badge-completed';
    if (status === 'معلق')  return 'badge-pending';
    if (status === 'ملغي' || status === 'مرفوض') return 'badge-error';
    return 'badge-available';
    // ملاحظة: أضف .badge-error في CSS إن لزم
  };

  const rows = items.map((i, idx) => `
    <tr>
      <td>${idx+1}</td>
      <td><strong>${i.workshopName || '-'}</strong></td>
      <td>${i.date || '-'}</td>
      <td>${i.registrationTime || '-'}</td>
      <td><span class="workshop-badge ${badge(i.status)}">${i.status || '-'}</span></td>
      <td>
        ${i.status === 'معلق'
          ? '<span style="color:var(--tvtc-text-muted)">بانتظار اعتماد المدرب</span>'
          : (i.status === 'معتمد'
              ? '<span style="color:var(--success)">تم الاعتماد ✓</span>'
              : '<span style="color:var(--tvtc-text-muted)">—</span>')}
      </td>
    </tr>
  `).join('');

  el.innerHTML = `
    ${isFallback ? '<div class="info" style="margin-bottom:8px;color:#92400e;background:#fffbeb;padding:10px;border-radius:8px;">عرض مبسّط بناءً على السجلات المعتمدة فقط.</div>' : ''}
    <table>
      <thead>
        <tr>
          <th>#</th><th>اسم الورشة</th><th>تاريخ الورشة</th><th>وقت التسجيل</th><th>الحالة</th><th>ملاحظات</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}



function renderTraineeStats(stats) {
  const html = `
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E6F1EC;color:var(--tvtc-primary);">⏱️</div></div>
      <div class="stat-value">${stats.totalHours}</div>
      <div class="stat-label">إجمالي الساعات المكتسبة</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#D1FAE5;color:var(--success);">✅</div></div>
      <div class="stat-value">${stats.completedWorkshops}</div>
      <div class="stat-label">الورش المكتملة</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#FEF3C7;color:var(--warning);">⏳</div></div>
      <div class="stat-value">${stats.pendingWorkshops}</div>
      <div class="stat-label">في انتظار الاعتماد</div>
    </div>`;
  document.getElementById('traineeStats').innerHTML = html;
}

function renderAvailableWorkshops(workshops) {
  const html = (workshops||[]).map(w => `
    <div class="workshop-card">
      <div class="workshop-header">
        <div class="workshop-title">${w.name}</div>
        <span class="workshop-badge badge-${w.status === 'متاح' ? 'available' : 'completed'}">${w.status}</span>
      </div>
      <div class="workshop-details">
        <div class="workshop-detail"><span>⏱️</span><span>${w.hours} ساعات</span></div>
        <div class="workshop-detail"><span>📅</span><span>${w.date}</span></div>
        <div class="workshop-detail"><span>📍</span><span>${w.location}</span></div>
        <div class="workshop-detail"><span>👥</span><span>${w.registered}/${w.capacity} مسجل</span></div>
      </div>
      <div class="workshop-actions">
        ${w.status === 'متاح'
          ? `<button class="btn btn-primary btn-small" onclick="registerWorkshop('${w.id}')">تسجيل حضور</button>`
          : `<button class="btn btn-outline btn-small" disabled>مكتمل</button>`}
        <button class="btn btn-outline btn-small" onclick="viewWorkshopDetails('${w.id}')">التفاصيل</button>
      </div>
    </div>`).join('');
  document.getElementById('availableWorkshops').innerHTML = html;
}

function renderTraineeSkills(skills) {
  const html = `
    <table>
      <thead><tr><th>المهارة</th><th>الساعات</th><th>التاريخ</th><th>الحالة</th><th>الشهادة</th></tr></thead>
      <tbody>
        ${(skills||[]).map(s => `
          <tr>
            <td><strong>${s.name}</strong></td>
            <td>${s.hours}</td>
            <td>${s.date}</td>
            <td><span class="workshop-badge badge-${s.status === 'معتمد' ? 'completed' : 'pending'}">${s.status}</span></td>
            <td>${s.status === 'معتمد'
                ? `<button class="btn btn-accent btn-small" onclick="downloadCertificate('${s.id}')">تحميل</button>`
                : `<button class="btn btn-outline btn-small" disabled>غير متاح</button>`}
            </td>
          </tr>`).join('')}
      </tbody>
    </table>`;
  document.getElementById('traineeSkillsTable').innerHTML = html;
}

async function registerWorkshop(workshopId) {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({ action:'registerWorkshop', userId: currentUser.id, workshopId })
    });
    const data = await res.json();
    if (data?.success) {
      alert('✅ تم تسجيل حضورك بنجاح!\nسيتم اعتماده من قبل المدرب.');
      await loadTraineeData();
    } else {
      alert('❌ ' + (data?.message || 'تعذر التسجيل'));
    }
  } catch (e) {
    console.error('خطأ في التسجيل:', e);
    alert('حدث خطأ. الرجاء المحاولة مرة أخرى');
  } finally {
    spinner?.classList.add('hidden');
  }
}

// ---- لوحة المدرب ----
async function showTrainerDashboard() {
  document.getElementById('trainerPage')?.classList.remove('hidden');
  document.getElementById('trainerInfo').textContent = currentUser.name;
  await loadTrainerData();
  await loadWorkshopTypes();
}

async function loadWorkshopTypes() {
  try {
    const res = await fetch(`${CONFIG.GOOGLE_SCRIPT_URL}?action=getWorkshopTypes`);
    const data = await res.json();
    if (data?.success) {
      const select = document.getElementById('workshopName');
      select.innerHTML = '<option value="">-- اختر من القائمة --</option>' +
        data.workshopTypes.map(t => `<option value="${t.name}" data-hours="${t.hours}">${t.name} (${t.hours} ساعات)</option>`).join('');
    }
  } catch (e) {
    console.error('خطأ في جلب أنواع الورش:', e);
  }
}

// إضافة ورشة (إرسال فقط؛ الإضافة تتم في GAS)
async function submitWorkshop(event) {
  event.preventDefault();

  const select = document.getElementById('workshopName');
  const location = document.getElementById('workshopLocation').value.trim();
  const dateStr  = document.getElementById('workshopDate').value.trim();
  const timeStr  = document.getElementById('workshopTime').value.trim();
  const capacity = Number(document.getElementById('workshopCapacity').value);

  // قراءة الساعات من data-hours أو من نص الخيار
  let hours = Number(select.selectedOptions[0]?.dataset.hours || 0);
  if (!hours) {
    const m = /\((\d+)\s*س/.exec(select.selectedOptions[0]?.textContent || '');
    if (m) hours = Number(m[1]);
  }

  // بدائل آمنة لو department/ trainerId فاضية من الحساب
  const trainerId  = String(currentUser?.id || '').trim();
  const department = String(currentUser?.department || '').trim();

  // تجميع النواقص لعرضها للمستخدم
  const missing = [];
  if (!select.value.trim()) missing.push('اسم الورشة');
  if (!hours)               missing.push('الساعات');
  if (!dateStr)             missing.push('التاريخ');
  if (!timeStr)             missing.push('الوقت');
  if (!location)            missing.push('المكان');
  if (!capacity)            missing.push('المقاعد');
  if (!department)          missing.push('القسم');
  if (!trainerId)           missing.push('المعرّف');

  if (missing.length) {
    alert('⚠️ تأكد من تعبئة: ' + missing.join('، '));
    console.warn('Missing fields:', missing);
    return;
  }

  // حمل الإرسال
  const workshopData = {
    action: 'addWorkshop',
    trainerId,
    name: select.value.trim(),
    location,
    date: dateStr,
    time: timeStr,
    hours,
    capacity,
    department,
    status: 'متاح'
  };

  console.log('payload to GAS:', workshopData);

  const spinner = document.getElementById('loadingSpinner');
  spinner.classList.remove('hidden');
  try {
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, { method:'POST', body: JSON.stringify(workshopData) });
    const data = await res.json();
    if (data.success) {
      alert('✅ تم نشر الورشة بنجاح!');
      closeModal();
      await loadTrainerData();
    } else {
      alert('❌ ' + (data.message || 'تعذر إضافة الورشة') + (data.missing ? '\nمفقود: ' + data.missing.join(', ') : ''));
    }
  } catch (err) {
    console.error('خطأ في إضافة الورشة:', err);
    alert('❌ حدث خطأ أثناء الإرسال.');
  } finally {
    spinner.classList.add('hidden');
  }
}


function parseDateFlexible_(s){
  if (!s) return null;
  const t = String(s).trim();

  // ISO: 2025-03-14
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
    const [y,m,d] = t.split('-').map(Number);
    return new Date(y, m-1, d);
  }

  // DD/MM/YYYY أو MM/DD/YYYY — نفك التباس بحسب قيمة >12
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(t)) {
    const [a,b,y] = t.split('/').map(Number);
    if (a > 12 && b <= 12) {           // واضح أنها DD/MM/YYYY
      return new Date(y, b-1, a);
    } else if (b > 12 && a <= 12) {    // واضح أنها MM/DD/YYYY
      return new Date(y, a-1, b);
    } else {
      // كلاهما <= 12 — اعتبرها DD/MM أولاً (سياسة عربية)
      return new Date(y, b-1, a);
    }
  }

  // محاولة أخيرة من المتصفح
  const dflt = new Date(t);
  return isNaN(dflt.getTime()) ? null : dflt;
}



function renderTrainerStats(stats) {
  const html = `
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E6F1EC;color:var(--tvtc-primary);">📚</div></div>
      <div class="stat-value">${stats.activeWorkshops}</div>
      <div class="stat-label">الورش النشطة</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E0E7FF;color:#4F46E5;">👥</div></div>
      <div class="stat-value">${stats.totalStudents}</div>
      <div class="stat-label">إجمالي المتدربين</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#FEF3C7;color:var(--warning);">⏳</div></div>
      <div class="stat-value">${stats.pendingApprovals}</div>
      <div class="stat-label">في انتظار الاعتماد</div>
    </div>`;
  document.getElementById('trainerStats').innerHTML = html;
}

function renderTrainerWorkshops(workshops) {
  const html = `
    <table>
      <thead><tr><th>اسم الورشة</th><th>التاريخ</th><th>المكان</th><th>المسجلين</th><th>الحالة</th><th>الإجراءات</th></tr></thead>
      <tbody>
        ${(workshops||[]).map(w => `
          <tr>
            <td><strong>${w.name}</strong></td>
            <td>${w.date}</td>
            <td>${w.location}</td>
            <td>${w.registered}/${w.capacity}</td>
            <td><span class="workshop-badge badge-${(w.status === 'نشط' || w.status === 'متاح') ? 'available' : 'completed'}">${w.status}</span></td>
            <td><button class="btn btn-primary btn-small" onclick="viewWorkshopDetails('${w.id}')">التفاصيل</button></td>
          </tr>`).join('')}
      </tbody>
    </table>`;
  document.getElementById('trainerWorkshops').innerHTML = html;
}

function renderPendingAttendance(attendance) {
  const html = `
    <table>
      <thead>
        <tr>
          <th>اسم المتدرب</th>
          <th>رقم التدريب</th>
          <th>الورشة</th>
          <th>وقت التسجيل</th>
          <th>الحالة</th>
          <th>الإجراء</th>
        </tr>
      </thead>
      <tbody>
        ${((attendance || []).length ? (attendance || []).map(a => {
          const badge = '<span class="workshop-badge badge-pending">معلق</span>';
          return `
            <tr>
              <td><strong>${a.traineeName}</strong></td>
              <td>${a.traineeId}</td>
              <td>${a.workshopName}</td>
              <td>${a.registrationTime}</td>
              <td>${badge}</td>
              <td>
                <select class="input input-small" onchange="setAttendanceStatusUI('${a.id}', this)">
                  <option value="pending" selected>معلق</option>
                  <option value="approved">معتمد (حضور)</option>
                  <option value="noshow">لم يحضر</option>
                </select>
              </td>
            </tr>
          `;
        }).join('') : `
          <tr>
            <td colspan="6" style="text-align:center;color:var(--tvtc-text-muted);">
              لا توجد سجلات معلّقة حاليًا.
            </td>
          </tr>
        `)}
      </tbody>
    </table>`;
  document.getElementById('pendingAttendance').innerHTML = html;
}


async function approveAttendance(attendanceId, button) {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method:'POST',
      body: JSON.stringify({ action:'approveAttendance', attendanceId, trainerId: currentUser.id })
    });
    const data = await res.json();
    if (data?.success) {
      const row = button.closest('tr');
      const badge = row.querySelector('.workshop-badge');
      badge.className = 'workshop-badge badge-completed';
      badge.textContent = 'معتمد';
      button.textContent = 'معتمد ✓';
      button.disabled = true;
      button.classList.remove('btn-primary');
      button.classList.add('btn-outline');
      alert('✅ تم اعتماد الحضور بنجاح!');
    } else {
      alert('❌ ' + (data?.message || 'تعذر الاعتماد'));
    }
  } catch (e) {
    console.error('خطأ في اعتماد الحضور:', e);
    alert('حدث خطأ. الرجاء المحاولة مرة أخرى');
  } finally {
    spinner?.classList.add('hidden');
  }
}

async function setAttendanceStatusUI(attendanceId, selectEl) {
  const newVal = selectEl.value; // 'approved' | 'pending' | 'noshow'
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method:'POST',
      body: JSON.stringify({ action:'setAttendanceStatus', attendanceId, trainerId: currentUser.id, status: newVal })
    });
    const data = await res.json();
    if (data?.success) {
      // حدّث الشارة في نفس الصف (إن وُجدت)
      const row = selectEl.closest('tr');
      const badge = row?.querySelector('.workshop-badge');
      if (badge) {
        if (newVal === 'approved') {
          badge.className = 'workshop-badge badge-completed';
          badge.textContent = 'معتمد';
        } else if (newVal === 'pending') {
          badge.className = 'workshop-badge badge-pending';
          badge.textContent = 'معلق';
        } else {
          badge.className = 'workshop-badge badge-error';
          badge.textContent = 'لم يحضر';
        }
      }
      // تحديثات اختيارية: أعد تحميل الكاش/القوائم لو حبيت
      // await loadTrainerData();
      try { await loadTrainerData(); } catch {}
      alert('✅ تم تحديث الحالة بنجاح.');
    } else {
      showError(data?.message || 'تعذر تحديث الحالة'); // نفس مكوّن الخطأ عندك أعلى الصفحة
    }
  } catch (e) {
    console.error('خطأ في تغيير الحالة:', e);
    alert('حدث خطأ أثناء تغيير الحالة');
  } finally {
    spinner?.classList.add('hidden');
  }
}


// ---- لوحات رئيس القسم / الإدارة ----
async function showHeadDashboard() {
  document.getElementById('headPage')?.classList.remove('hidden');
  document.getElementById('headInfo').textContent = `${currentUser.name} - ${currentUser.department}`;
  await loadHeadData();
}

async function showAdminDashboard() {
  document.getElementById('adminPage')?.classList.remove('hidden');
  document.getElementById('adminInfo').textContent = currentUser.name;
  await loadAdminData();
}

async function loadHeadData() {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getHeadData&userId=${encodeURIComponent(currentUser.id)}&department=${encodeURIComponent(currentUser.department)}`;
    const res = await fetch(url);
    const data = await res.json();
    if (data?.success) {
      renderHeadStats(data.stats);
      renderDepartmentReports(data.reports);
      renderTopTrainees(data.topTrainees);
    }
  } catch (e) {
    console.error('خطأ في جلب البيانات:', e);
  } finally {
    spinner?.classList.add('hidden');
  }
}

function renderHeadStats(stats) {
  const html = `
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E6F1EC;color:var(--tvtc-primary);">📚</div></div>
      <div class="stat-value">${stats.totalWorkshops || 0}</div>
      <div class="stat-label">إجمالي الورش</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E0E7FF;color:#4F46E5;">👥</div></div>
      <div class="stat-value">${stats.totalStudents || 0}</div>
      <div class="stat-label">عدد المتدربين</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#D1FAE5;color:var(--success);">📈</div></div>
      <div class="stat-value">${stats.averageAttendance || 0}%</div>
      <div class="stat-label">نسبة الحضور</div>
    </div>`;
  document.getElementById('headStats').innerHTML = html;
}

function renderDepartmentReports(reports) {
  const html = `
    <table>
      <thead><tr><th>الورشة</th><th>المدرب</th><th>عدد الحضور</th><th>نسبة الحضور</th><th>التاريخ</th><th>الحالة</th></tr></thead>
      <tbody>
        ${(reports&&reports.length)? reports.map(r => `
          <tr>
            <td><strong>${r.workshopName}</strong></td>
            <td>${r.trainerName}</td>
            <td>${r.attendance}/${r.capacity}</td>
            <td>${r.attendanceRate}%</td>
            <td>${r.date}</td>
            <td><span class="workshop-badge badge-${r.status === 'نشط' ? 'available' : 'completed'}">${r.status}</span></td>
          </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;">لا توجد بيانات</td></tr>'}
      </tbody>
    </table>`;
  document.getElementById('departmentReports').innerHTML = html;
}

function renderTopTrainees(trainees) {
  const icons = ['🥇','🥈','🥉'];
  const colors = ['#FFD700','#C0C0C0','#CD7F32'];
  const html = (trainees&&trainees.length)? trainees.slice(0,3).map((t,i)=>`
    <div class="workshop-card" style="border-right-color:${colors[i]};">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div style="font-size:24px;margin-bottom:5px;">${icons[i]}</div>
          <div class="workshop-title">${t.name}</div>
          <div style="color:var(--tvtc-text-muted);font-size:14px;margin-top:5px;">${t.totalHours} ساعة تدريبية</div>
        </div>
      </div>
    </div>`).join('') : '<p style="text-align:center;color:var(--tvtc-text-muted);">لا توجد بيانات</p>';
  document.getElementById('topTrainees').innerHTML = html;
}

// ---- لوحة الإدارة ----
async function loadAdminData() {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const res = await fetch(`${CONFIG.GOOGLE_SCRIPT_URL}?action=getAdminData`);
    const data = await res.json();
    if (data?.success) {
      renderAdminStats(data.stats);
      renderAllDepartmentsReports(data.departments);
    }
  } catch (e) {
    console.error('خطأ في جلب البيانات:', e);
  } finally {
    spinner?.classList.add('hidden');
  }
}

function renderAdminStats(stats) {
  const html = `
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E6F1EC;color:var(--tvtc-primary);">🏢</div></div>
      <div class="stat-value">${stats.totalDepartments || 0}</div>
      <div class="stat-label">عدد الأقسام</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#E0E7FF;color:#4F46E5;">📚</div></div>
      <div class="stat-value">${stats.totalWorkshops || 0}</div>
      <div class="stat-label">إجمالي الورش</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#FEF3C7;color:var(--warning);">👥</div></div>
      <div class="stat-value">${stats.totalStudents || 0}</div>
      <div class="stat-label">إجمالي المتدربين</div>
    </div>
    <div class="stat-card">
      <div class="stat-card-header"><div class="stat-icon" style="background:#D1FAE5;color:var(--success);">📊</div></div>
      <div class="stat-value">${stats.averageSuccess || 0}%</div>
      <div class="stat-label">نسبة النجاح العامة</div>
    </div>`;
  document.getElementById('adminStats').innerHTML = html;
}

function renderAllDepartmentsReports(departments) {
  const html = `
    <table>
      <thead><tr><th>القسم</th><th>عدد الورش</th><th>عدد المتدربين</th><th>نسبة الحضور</th><th>الأداء</th><th>التقرير</th></tr></thead>
      <tbody>
        ${(departments&&departments.length)? departments.map(d => `
          <tr>
            <td><strong>${d.name}</strong></td>
            <td>${d.totalWorkshops}</td>
            <td>${d.totalStudents}</td>
            <td>${d.attendanceRate}%</td>
            <td><span class="workshop-badge badge-${d.performance === 'ممتاز' ? 'completed' : 'available'}">${d.performance}</span></td>
            <td><button class="btn btn-accent btn-small" onclick="viewDepartmentDetails('${d.name}')">عرض</button></td>
          </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;">لا توجد بيانات</td></tr>'}
      </tbody>
    </table>`;
  document.getElementById('allDepartmentsReports').innerHTML = html;
}

// ✅ توحيد نص عربي بسيط للمقارنات (يضمن مطابقة الأنواع حتى لو فيها مسافات/حروف مختلفة)
function _norm(s){
  return String(s||'').trim().replace(/\s+/g,' ').replace(/[ـ]/g,'').toLowerCase();
}


function _pickOfficialsFromUsers(users, traineeDept){
  const N = s => String(s||'')
    .trim()
    .replace(/[ـ\u200f\u200e]/g,'')
    .replace(/\s+/g,' ')
    .replace(/^قسم\s+/i,'')
    .replace(/\((?:بنين|بنات|طلاب|طالبات|شطر.*?)\)/gi,'')
    .toLowerCase();

  const eqLoose = (a,b) => {
    const A=N(a), B=N(b);
    if (!A || !B) return false;
    return A===B || A.includes(B) || B.includes(A);
  };

  // حاول نستخرج القسم من نص الدور إذا عمود القسم فاضي
  const deptFromRole = (roleRaw) => {
    const m = String(roleRaw||'').match(/رئيس\s*قسم\s*[:\-–—]?\s*(.+)$/i);
    return m ? m[1].trim() : '';
  };

  const traineeDeptN = N(traineeDept);
  let hodExact='', hodAny='', deanStd='', dean='';

  for (const u of (users||[])) {
    const roleRaw = String(u.userType ?? u.UserType ?? u.role ?? '').trim();
    const depRaw  = String(u.department ?? u.Department ?? '').trim()
                  || deptFromRole(roleRaw); // 👈 جديد: استخرج القسم من الدور
    const name    = String(u.name ?? u.Name ?? u.displayName ?? '').trim();
    if (!name) continue;

    const isHOD   = /رئيس\s*قسم/i.test(roleRaw);
    const isDeanS = /وكيل\s*شؤون\s*المتدربين/i.test(roleRaw);
    const isDean  = /(^|\s)عميد(\s|$)/i.test(roleRaw) && !/شؤون\s*(?:الطلاب|المتدربين)/i.test(roleRaw);

    if (isHOD) {
      if (!hodAny) hodAny = name;
      if (!hodExact && traineeDeptN && (eqLoose(depRaw, traineeDeptN))) hodExact = name;
    }
    if (!deanStd && isDeanS) deanStd = name;
    if (!dean && isDean)     dean    = name;

    if (hodExact && deanStd && dean) break;
  }

  return { hod: (hodExact || hodAny || ''), deanStd, dean };
}


// ✅ محاولة إحضار المسؤولين من GAS
async function _fetchOfficialsFromGAS(department){
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getOfficials&department=${encodeURIComponent(department||'')}`;
    const r = await fetch(url);
    const j = await r.json();

    if (j?.success) {
      // الشكل المتوقّع: {success:true, hod:'...', deanOfStudents:'...', dean:'...'}
      return {
        hod: j.hod || j.headOfDepartment || '',
        deanStd: j.deanOfStudents || j.deanStd || '',
        dean: j.dean || ''
      };
    }
    if (j && j.message === 'Invalid action') {
      // سطر خطأ:
      console.error('❌ getOfficials غير متاح على GAS (Invalid action)'); // ← سطر الخطأ
      // سطر تصحيح: جرّب إحضار قائمة المستخدمين ثم انتقِ المطلوبين
      const urlUsers = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getUsers`;
      const r2 = await fetch(urlUsers);
      const j2 = await r2.json();
      if (Array.isArray(j2?.users)) {
        const picked = _pickOfficialsFromUsers(j2.users, department);
        return { hod: picked.hod, deanStd: picked.deanStd, dean: picked.dean };
      }
      return { hod:'', deanStd:'', dean:'' };
    }

    // استجابة غير متوقعة
    console.error('⚠️ استجابة غير متوقعة من getOfficials:', j); // ← سطر الخطأ
    return { hod:'', deanStd:'', dean:'' };
  } catch (e) {
    console.error('❌ خطأ شبكة أثناء جلب المسؤولين:', e); // ← سطر الخطأ
    return { hod:'', deanStd:'', dean:'' };
  }
}

// ===========================
// 🔁 تطبيع ومطابقة عربية قوية
// ===========================
function _anorm(s){
  return String(s||'')
    .trim()
    .replace(/[ـ]/g,'')   // إزالة التطويل
    .replace(/\s+/g,' ')  // توحيد المسافات
    .toLowerCase();
}
function aeq(a,b){ return _anorm(a) === _anorm(b); }

// نقرأ الدور من كل المفاتيح المحتملة
function getRole(u){
  return String(u?.userType ?? u?.UserType ?? u?.role ?? '').trim();
}

// كواشف الأدوار (مرنة)
// كواشف الأدوار (مرنة)
function isHOD(u){     return /رئيس\s*قسم/i.test(getRole(u)); }
function isDeanStd(u){ return /وكيل\s*شؤون\s*المتدربين/i.test(getRole(u)); }
function isDean(u){    return /عميد/i.test(getRole(u)); }
// ===========================

// ✅ جلب دليل المستخدمين مع محاولات fallback متعددة

async function fetchUsersDirectory() {
  const tryActions = ['getUsersDirectory', 'getAllUsers', 'listUsers'];
  for (const action of tryActions) {
    try {
      const res = await fetch(`${CONFIG.GOOGLE_SCRIPT_URL}?action=${action}`);
      const data = await res.json();
      // نتوقع شكل: { success:true, users:[{name, userType, department}] }
      if (data?.success && Array.isArray(data.users)) return data.users;
    } catch (e) {
      console.warn('fetchUsersDirectory fallback:', action, e);
    }
  }
  console.error('❌ لم أتمكن من جلب دليل المستخدمين من GAS'); // سطر خطأ
  return [];
}


// ============================================
// تصدير سجل المهارات من قالب محلي باستخدام ExcelJS
// يعتمد على CONFIG.EXPORT في config.js
// ============================================
async function exportTraineeExcel() {
  try {
    if (!currentUser || !currentUser.id) {
      showError('يجب تسجيل الدخول أولًا');
      return;
    }

    showFileOverlay('🔄 تهيئة البيانات من الخادم...'); // سطر تصحيح: بدء التحميل المرئي

    // 1) اطلب بيانات التصدير من GAS
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=exportTraineeXlsx&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    const j = await res.json();
    if (!j.success) {
      console.error('❌ exportTraineeExcel: GAS error:', j.message); // سطر خطأ
      showError(j.message || 'تعذر تجهيز بيانات التصدير');
      await hideFileOverlay(); // سطر تصحيح
      return;
    }

    const td = j.traineeData || {};
    const skillsArr = j.skills || [];

    // 2) حمّل القالب
    updateFileOverlay('⬇️ تحميل القالب المحلي...'); // سطر تصحيح
    const tplRes = await fetch(CONFIG.EXPORT.TEMPLATE_URL, { cache: 'no-cache' });
    if (!tplRes.ok) {
      console.error('❌ exportTraineeExcel: template fetch failed:', tplRes.status, tplRes.statusText); // سطر خطأ
      showError('تعذر تحميل قالب Excel المحلي');
      await hideFileOverlay(); // سطر تصحيح
      return;
    }
    const tplBuf = await tplRes.arrayBuffer();

    // 3) افتح القالب واكتب البيانات
    updateFileOverlay('📝 كتابة البيانات داخل القالب...'); // سطر تصحيح
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(tplBuf);
    const ws = wb.getWorksheet(CONFIG.EXPORT.SHEET_NAME);
    if (!ws) {
      console.error('❌ exportTraineeExcel: sheet not found:', CONFIG.EXPORT.SHEET_NAME); // سطر خطأ
      showError('ورقة القالب غير موجودة داخل الملف');
      await hideFileOverlay(); // سطر تصحيح
      return;
    }

    // رؤوس المتدرب
    ws.getCell(CONFIG.EXPORT.NAME_CELL).value  = td.name      || '';
    ws.getCell(CONFIG.EXPORT.ID_CELL).value    = td.studentId || '';
    ws.getCell(CONFIG.EXPORT.MAJOR_CELL).value = td.major     || '';
    ws.getCell(CONFIG.EXPORT.SEM_CELL).value   = td.semester  || '';
    ws.getCell(CONFIG.EXPORT.YEAR_CELL).value  = td.year      || '';

    // التواقيع
    // جهّز القسم لو جاء باسم مختلف
    td.major = td.major || td.department || td.dept || currentUser?.department || '';
    // التواقيع (هجينة: GAS getOfficials → getUsers → Excel محلي)
    await resolveSignaturesOnline(td);
    ws.getCell(CONFIG.EXPORT.HOD_CELL)      .value = td.headOfDepartment || '';
    ws.getCell(CONFIG.EXPORT.DEAN_STD_CELL) .value = td.deanOfStudents   || '';
    ws.getCell(CONFIG.EXPORT.DEAN_CELL)     .value = td.dean             || '';


    // المهارات
    const MAP = CONFIG.EXPORT.SKILLS_MAP || {};
    let total = 0;
    for (const { name, hours } of skillsArr) {
      const cellAddr = MAP[name];
      if (!cellAddr) continue;
      ws.getCell(cellAddr).value = Number(hours) || 0;
      const rowNum = Number(cellAddr.replace(/[A-Z]/gi,'')); // اسم المهارة في العمود B (اختياري)
      if (!isNaN(rowNum)) {
        const nameCell = 'B' + rowNum;
        if (!ws.getCell(nameCell).value) ws.getCell(nameCell).value = name;
      }
      total += Number(hours) || 0;
    }
    ws.getCell(CONFIG.EXPORT.TOTAL_HOURS_CELL).value = total;

    // 4) التحضير للتنزيل
    updateFileOverlay('📦 تجهيز الملف للتنزيل...'); // سطر تصحيح
    const safeId = (td.studentId || currentUser.id || '').toString().trim();
    const outName = `سجل_المهارات_${safeId}.xlsx`;
    const outBuf = await wb.xlsx.writeBuffer();

    // تنزيل
    const blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = outName;
    document.body.appendChild(a);
    a.click();
    URL.revokeObjectURL(a.href);
    a.remove();

    updateFileOverlay('✅ تم إنشاء الملف.. يبدأ التنزيل الآن'); // سطر تصحيح
  } catch (err) {
    console.error('❌ exportTraineeExcel crashed:', err); // سطر خطأ
    showError('حدث خطأ أثناء إنشاء ملف Excel');
  } finally {
    await hideFileOverlay(1200); // سطر تصحيح: إبقاء الأوفرلاي ظاهر ≥ 1.2 ثانية
  }
}


// ================================
//  قرّاءه المسؤولين من Excel محليًا (ExcelJS)
// ================================
function _guessHeaderMap(headerRow) {
  // محاولة ذكية لمطابقة رؤوس الأعمدة حتى لو كانت بالعربي/إنجليزي
  const N = (s)=>String(s||'').trim().toLowerCase().replace(/[ـ]/g,'').replace(/\s+/g,' ');
  const map = { role: null, dept: null, name: null };

  headerRow.forEach((h, idx) => {
    const hh = N(h);
    if (!map.role && /(role|دور|منصب|الصفة|الوظيفة|المسمى)/.test(hh)) map.role = idx;
    else if (!map.dept && /(dept|department|قسم|القسم|تخصص|التخصص)/.test(hh)) map.dept = idx;
    else if (!map.name && /(name|الاسم|اسم)/.test(hh)) map.name = idx;
  });

  return map;
}


async function _loadOfficialsFromExcel(url, sheetName) {
  // سطر خطأ:
  if (typeof ExcelJS === 'undefined') {
    console.error('❌ ExcelJS غير متوفر لقراءة ملف المسؤولين'); 
    return [];
  }

  // سطر تصحيح: فحص وبناء مسار نهائي مثل resolveTemplateUrl
  try {
    if (url.startsWith('/')) url = url.slice(1);
    url = url.replace(/\\/g,'/');
    const abs = new URL(url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now(), window.location.href);

    const resp = await fetch(abs.href, { cache: 'no-cache' });
    if (!resp.ok) {
      console.error('❌ فشل تحميل ملف المسؤولين من', abs.href, 'status=', resp.status);
      return [];
    }
    const buf = await resp.arrayBuffer();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);

    const ws = sheetName ? wb.getWorksheet(sheetName) : wb.worksheets[0];
    if (!ws) {
      console.error('❌ لم يتم العثور على الورقة المطلوبة لملف المسؤولين:', sheetName);
      return [];
    }

    // نفترض الصف الأول رؤوس
    const headerRow = ws.getRow(1).values.slice(1); // values[0] فارغة عادةً
    const map = _guessHeaderMap(headerRow);

    if (map.role == null || map.name == null) {
      console.error('❌ تعذّر اكتشاف أعمدة الدور/الاسم في ملف المسؤولين.');
      return [];
    }

    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const vals = ws.getRow(r).values.slice(1);
      const role = vals[map.role] ?? '';
      const dept = map.dept != null ? (vals[map.dept] ?? '') : '';
      const name = vals[map.name] ?? '';
      if (String(name).trim()) {
        rows.push({ role, department: dept, name });
      }
    }
    return rows;
  } catch (e) {
    console.error('❌ خطأ أثناء قراءة ملف المسؤولين Excel:', e);
    return [];
  }
}

function _pickOfficialsFromExcelRows(rows, traineeDept) {
  const N = (s)=>String(s||'').trim().toLowerCase().replace(/[ـ]/g,'').replace(/\s+/g,' ');
  const deptN = N(traineeDept);

  let hod='', deanStd='', dean='';

  for (const r of rows) {
    const role = N(r.role);
    const dep  = N(r.department);
    const name = r.name || '';

    // رئيس القسم يطابق القسم
    if (!hod && /(رئيس\s*قسم|head\s*of\s*department|hod)/.test(role) && (dep ? dep===deptN : true)) {
      hod = name;
    }
    // وكيل شؤون المتدربين بدون اشتراط قسم
    if (!deanStd && /(وكيل\s*شؤون\s*المتدربين|dean\s*of\s*students|student\s*affairs)/.test(role)) {
      deanStd = name;
    }
    // العميد
    if (!dean && /(عميد|dean(?!.*students))/.test(role)) {
      dean = name;
    }

    if (hod && deanStd && dean) break;
  }

  return { hod, deanStd, dean };
}

// ===========================
// تطبيع/مقارنات عربية مرنة
// ===========================
function _anorm(s){
  return String(s||'').trim().replace(/[ـ]/g,'').replace(/\s+/g,' ').toLowerCase();
}
function aeq(a,b){ return _anorm(a) === _anorm(b); }

// كشف الأدوار بالعربي
function _isHOD(role){     return /رئيس\s*قسم/i.test(role); }
function _isDeanStd(role){ return /وكيل\s*شؤون\s*المتدربين/i.test(role); }
function _isDean(role){    return /عميد/i.test(role); }


// بلايسهولدرات التوقيع
const ROLE_PLACEHOLDER_RE = /(رئيس\s*(?:ال)?قسم|وكيل\s*شؤون\s*المتدربين|عميد|العميد|^[-–—]+$)/i;

function _isPlaceholder(v){
  const s = String(v||'').trim();
  return !s || ROLE_PLACEHOLDER_RE.test(s); // فراغ أو الكلمة المجردة فقط
}
function _cleanIfPlaceholder(v){ return _isPlaceholder(v) ? '' : v; }


// ================================
// ✅ جلب المسؤولين أونلاين فقط من نفس Google Sheet (GAS)
// الأكشن: getUsers  => {success:true, users:[{name, userType, department}]}
// نسخة ألين: تفضّل نفس القسم، وإن ما لقت تطابقاً كاملاً تجيب أي رئيس قسم

async function _getOfficialsOnlineStrict(department){
  const traineeDept = String(department||'').trim();

  const N = s => String(s||'')
    .trim()
    .replace(/[ـ\u200f\u200e]/g,'')
    .replace(/\s+/g,' ')
    .replace(/^قسم\s+/i,'')
    .replace(/\((?:بنين|بنات|طلاب|طالبات|شطر.*?)\)/gi,'')
    .toLowerCase();
  const eqLoose = (a,b) => {
    const A=N(a), B=N(b);
    if (!A || !B) return false;
    return A===B || A.includes(B) || B.includes(A);
  };
  const deptFromRole = (roleRaw) => {
    const m = String(roleRaw||'').match(/رئيس\s*قسم\s*[:\-–—]?\s*(.+)$/i);
    return m ? m[1].trim() : '';
  };

  try {
    const r = await fetch(`${CONFIG.GOOGLE_SCRIPT_URL}?action=getUsers`);
    const j = await r.json();
    if (!j?.success || !Array.isArray(j.users)) return { hod:'', deanStd:'', dean:'' };

    let hodExact='', hodAny='', deanStd='', dean='';
    for (const u of j.users) {
      const roleRaw = String(u.userType ?? u.UserType ?? u.role ?? '').trim();
      const depRaw  = String(u.department ?? u.Department ?? '').trim() || deptFromRole(roleRaw);
      const name    = String(u.name ?? u.Name ?? u.displayName ?? '').trim();
      if (!name) continue;

      const isHOD   = /رئيس\s*قسم/i.test(roleRaw);
      const isDeanS = /وكيل\s*شؤون\s*المتدربين/i.test(roleRaw);
      const isDean_ = /(^|\s)عميد(\s|$)/i.test(roleRaw) && !/شؤون\s*(?:الطلاب|المتدربين)/i.test(roleRaw);

      if (isHOD) {
        if (!hodAny) hodAny = name;
        if (!hodExact && traineeDept && eqLoose(depRaw, traineeDept)) hodExact = name;
      }
      if (!deanStd && isDeanS) deanStd = name;
      if (!dean && isDean_)    dean    = name;

      if (hodExact && deanStd && dean) break;
    }
    return { hod: (hodExact || hodAny || ''), deanStd, dean };
  } catch {
    return { hod:'', deanStd:'', dean:'' };
  }
}


// ================================
// ✅ الدالة الرئيسية (صارمة وأونلاين فقط)
async function resolveSignaturesOnlineStrict(traineeData){
  // تنظيف أي قيمة Placeholder
  traineeData.headOfDepartment = _cleanIfPlaceholder(traineeData.headOfDepartment);
  traineeData.deanOfStudents   = _cleanIfPlaceholder(traineeData.deanOfStudents);
  traineeData.dean             = _cleanIfPlaceholder(traineeData.dean);

  const alreadyDone =
    !_isPlaceholder(traineeData.headOfDepartment) &&
    !_isPlaceholder(traineeData.deanOfStudents) &&
    !_isPlaceholder(traineeData.dean);
  if (alreadyDone) return;

  // التقط القسم من عدة حقول (أضفنا major)
  const dept =
    traineeData.department ??
    traineeData.major ??            // ← مهم
    traineeData.dept ??
    traineeData.majorDepartment ??
    traineeData.departmentName ?? '';

  // استدعِ البحث الصارم
  const picked = await _getOfficialsOnlineStrict(dept);

  // املأ فقط الفارغ/البلايسهولدر
  if (_isPlaceholder(traineeData.headOfDepartment) && picked.hod)     traineeData.headOfDepartment = picked.hod;
  if (_isPlaceholder(traineeData.deanOfStudents)   && picked.deanStd) traineeData.deanOfStudents   = picked.deanStd;
  if (_isPlaceholder(traineeData.dean)             && picked.dean)    traineeData.dean             = picked.dean;

  const done =
    !_isPlaceholder(traineeData.headOfDepartment) &&
    !_isPlaceholder(traineeData.deanOfStudents) &&
    !_isPlaceholder(traineeData.dean);

  if (!done) {
    console.warn('⚠️ لم يتم العثور على جميع التواقيع (وفق التطابق الحرفي). تحقق من UserType والقسم في الشيت.', {
      dept,
      HOD: traineeData.headOfDepartment || '(غير موجود)',
      DeanStd: traineeData.deanOfStudents || '(غير موجود)',
      Dean: traineeData.dean || '(غير موجود)'
    });
  }
}


// محاولة أولى: GAS getOfficials (إن توفّر)
async function _fetchOfficialsFromGAS_Department(department){
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getOfficials&department=${encodeURIComponent(department||'')}`;
    const r = await fetch(url);
    const j = await r.json();
    if (j?.success) {
      return {
        hod: j.hod || j.headOfDepartment || '',
        deanStd: j.deanOfStudents || j.deanStd || '',
        dean: j.dean || ''
      };
    }
    if (j && j.message === 'Invalid action') {
      console.error('❌ getOfficials غير متاح على GAS (Invalid action)'); // سطر خطأ
      return null;
    }
    console.warn('⚠️ استجابة غير متوقعة من getOfficials:', j);
    return null;
  } catch (e) {
    console.error('❌ خطأ شبكة أثناء getOfficials:', e); // سطر خطأ
    return null;
  }
}

// محاولة ثانية: GAS getUsers ثم الانتقاء حسب القسم/الدور
async function _fetchOfficialsFromUsersSheet(department){
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getUsers`;
    const r = await fetch(url);
    const j = await r.json();
    if (!j?.success || !Array.isArray(j.users)) {
      console.error('❌ لم أستطع جلب قائمة المستخدمين من GAS');
      return { hod:'', deanStd:'', dean:'' };
    }

    const dept = String(department || '').trim();
    let hod='', deanStd='', dean='';
    for (const u of j.users) {
      const name = u.name || u.Name || '';
      const role = String(u.userType || u.role || u.UserType || '').trim();
      const dep  = String(u.department || u.Department || '').trim();

      if (!hod && role === 'رئيس قسم' && dep === dept) hod = name;
      if (!deanStd && role === 'وكيل شؤون المتدربين') deanStd = name;
      if (!dean && role === 'عميد') dean = name;

      if (hod && deanStd && dean) break;
    }
    return { hod, deanStd, dean };
  } catch (e) {
    console.error('❌ خطأ شبكة أثناء getUsers:', e);
    return { hod:'', deanStd:'', dean:'' };
  }
}

// === الدالة الرئيسية: أونلاين فقط (GAS) ===
// === الدالة الرئيسية: هجين (GAS أولاً ثم Excel محلي) ===
async function resolveSignaturesOnline(traineeData){
  // تنظيف القيم placeholder
  traineeData.headOfDepartment = _cleanIfPlaceholder(traineeData.headOfDepartment);
  traineeData.deanOfStudents   = _cleanIfPlaceholder(traineeData.deanOfStudents);
  traineeData.dean             = _cleanIfPlaceholder(traineeData.dean);

  const alreadyDone =
    !_isPlaceholder(traineeData.headOfDepartment) &&
    !_isPlaceholder(traineeData.deanOfStudents) &&
    !_isPlaceholder(traineeData.dean);
  if (alreadyDone) return;

  const dept = traineeData.department || traineeData.major || traineeData.dept || traineeData.departmentName || '';

  // (1) جرّب GAS مباشرة
  let picked = await _fetchOfficialsFromGAS_Department(dept);
  if (picked) {
    if (_isPlaceholder(traineeData.headOfDepartment) && picked.hod)     traineeData.headOfDepartment = picked.hod;
    if (_isPlaceholder(traineeData.deanOfStudents)   && picked.deanStd) traineeData.deanOfStudents   = picked.deanStd;
    if (_isPlaceholder(traineeData.dean)             && picked.dean)    traineeData.dean             = picked.dean;
  }

  // تحقق بعد GAS
  const doneAfterGAS =
    !_isPlaceholder(traineeData.headOfDepartment) &&
    !_isPlaceholder(traineeData.deanOfStudents) &&
    !_isPlaceholder(traineeData.dean);
  if (doneAfterGAS) return;

  // (2) جرّب GAS getUsers كـ fallback
  picked = await _fetchOfficialsFromUsersSheet(dept);
  if (_isPlaceholder(traineeData.headOfDepartment) && picked.hod)     traineeData.headOfDepartment = picked.hod;
  if (_isPlaceholder(traineeData.deanOfStudents)   && picked.deanStd) traineeData.deanOfStudents   = picked.deanStd;
  if (_isPlaceholder(traineeData.dean)             && picked.dean)    traineeData.dean             = picked.dean;

  const doneAfterUsers =
    !_isPlaceholder(traineeData.headOfDepartment) &&
    !_isPlaceholder(traineeData.deanOfStudents) &&
    !_isPlaceholder(traineeData.dean);
  if (doneAfterUsers) return;

  // (3) Excel محلي كحل أخير
  try {
    const officialsUrl   = CONFIG?.EXPORT?.OFFICIALS_URL   || 'templates/officials.xlsx';
    const officialsSheet = CONFIG?.EXPORT?.OFFICIALS_SHEET || '';
    const rows = await _loadOfficialsFromExcel(officialsUrl, officialsSheet);
    const pick = _pickOfficialsFromExcelRows(rows, dept);

    if (_isPlaceholder(traineeData.headOfDepartment) && pick.hod)     traineeData.headOfDepartment = pick.hod;
    if (_isPlaceholder(traineeData.deanOfStudents)   && pick.deanStd) traineeData.deanOfStudents   = pick.deanStd;
    if (_isPlaceholder(traineeData.dean)             && pick.dean)    traineeData.dean             = pick.dean;

    if (_isPlaceholder(traineeData.headOfDepartment) ||
        _isPlaceholder(traineeData.deanOfStudents)   ||
        _isPlaceholder(traineeData.dean)) {
      console.warn('⚠️ لم يتم العثور على بعض التواقيع بعد كل المحاولات (GAS ثم Excel).', {
        dept, HOD: traineeData.headOfDepartment || '(غير موجود)',
        DeanStd: traineeData.deanOfStudents     || '(غير موجود)',
        Dean: traineeData.dean                  || '(غير موجود)'
      });
    }
  } catch (err) {
    console.error('❌ فشل fallback Excel لقراءة المسؤولين:', err);
  }
}

function disableWhileBusy(btn, busy=true) {
  if (!btn) return;
  if (busy) {
    btn.disabled = true;
    btn.dataset.prevText = btn.textContent;
    btn.textContent = '...جاري التحضير';
  } else {
    btn.disabled = false;
    if (btn.dataset.prevText) btn.textContent = btn.dataset.prevText;
  }
}

async function onClickExport(btn){
  try {
    disableWhileBusy(btn, true);
    await exportTraineeExcel();
  } finally {
    disableWhileBusy(btn, false);
  }
}


// ================================================
// تصدير Excel من القالب مع الحفاظ على الألوان والدمجات (ExcelJS)
// ================================================
async function exportTraineeExcelUsingTemplate(traineeData) {
  showFileOverlay('🔄 تهيئة البيانات...'); // سطر تصحيح

  try {
    if (typeof ExcelJS === 'undefined') {
      throw new Error('ExcelJS غير محمّل. أضف:\n<script src="https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js"></script>');
    }
    if (!traineeData || !traineeData.name) {
      throw new Error('بيانات المتدرب غير مكتملة.');
    }

    updateFileOverlay('⬇️ تحميل القالب...'); // سطر تصحيح
    const tplUrl = resolveTemplateUrl();
    if (!tplUrl) throw new Error('تعذر تحديد مسار القالب.');
    const resp = await fetch(tplUrl, { method: 'GET', cache: 'no-cache' });
    if (!resp.ok) throw new Error(`فشل تحميل القالب (${resp.status} ${resp.statusText})`);
    const arrayBuffer = await resp.arrayBuffer();

    updateFileOverlay('📝 كتابة البيانات داخل القالب...'); // سطر تصحيح
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    const sheetName = (CONFIG?.EXPORT?.SHEET_NAME || '').trim();
    const worksheet = (sheetName && workbook.getWorksheet(sheetName)) || workbook.worksheets[0];
    if (!worksheet) throw new Error('لم يتم العثور على ورقة في القالب.');

    const setCell = (addr, val) => {
      const cell = worksheet.getCell(addr);
      cell.value = val == null ? '' : val;
    };

    // رؤوس
    setCell(CONFIG.EXPORT.NAME_CELL,  traineeData.name || '');
    setCell(CONFIG.EXPORT.ID_CELL,    traineeData.studentId || '');
    setCell(CONFIG.EXPORT.MAJOR_CELL, traineeData.major || '');
    setCell(CONFIG.EXPORT.SEM_CELL,   traineeData.semester || '');
    setCell(CONFIG.EXPORT.YEAR_CELL,  traineeData.year || '');

    // التواقيع
    console.error('📌 قبل حل التواقيع:', {  // سطر خطأ (تتبّع)
      dept: traineeData.department || traineeData.major || '',
      HOD: traineeData.headOfDepartment, DeanStd: traineeData.deanOfStudents, Dean: traineeData.dean
    });
    // تأكد أن الـ major متوفر قبل حل التواقيع
    traineeData.major = traineeData.major || traineeData.department || traineeData.dept || currentUser?.department || '';
    await resolveSignaturesOnline(traineeData);

    setCell(CONFIG.EXPORT.HOD_CELL,      traineeData.headOfDepartment || '');
    setCell(CONFIG.EXPORT.DEAN_STD_CELL, traineeData.deanOfStudents   || '');
    setCell(CONFIG.EXPORT.DEAN_CELL,     traineeData.dean             || '');

    // المهارات
    const map = CONFIG?.EXPORT?.SKILLS_MAP || {};
    let totalHours = 0;

    Object.values(map).forEach(cellAddr => {
      const row = parseInt(String(cellAddr).replace(/[^0-9]/g,''),10);
      if (!isNaN(row)) setCell(`B${row}`, '');
    });

    const approved = Array.isArray(traineeData.skills)
      ? traineeData.skills.filter(s => Number(s?.hours)>0 && String(s?.status).trim()==='معتمد')
      : Object.entries(traineeData.skills || {}).map(([k,v]) => ({ name:k, hours:Number(v||0), status:'معتمد' }));

    if (!approved.length) throw new Error('لا توجد مهارات معتمدة للتصدير.');

    for (const skill of approved) {
      const hours = Number(skill.hours || 0);
      totalHours += hours;
      const cellAddr = map[skill.name];
      if (cellAddr) {
        setCell(cellAddr, hours);
        const row = parseInt(String(cellAddr).replace(/[^0-9]/g,''),10);
        if (!isNaN(row)) setCell(`B${row}`, skill.name);
      }
    }
    setCell(CONFIG.EXPORT.TOTAL_HOURS_CELL || 'F30', totalHours);

    updateFileOverlay('📦 تجهيز الملف للتنزيل...'); // سطر تصحيح
    const safeName = String(traineeData.name || 'متدرب').replace(/[\\/:*?"<>|]/g,'_').slice(0,50);
    const safeId   = String(traineeData.studentId || '').replace(/[\\/:*?"<>|]/g,'_');
    const filename = `سجل_المهارات_${safeName}_${safeId}.xlsx`;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);

    updateFileOverlay('✅ تم إنشاء الملف.. يبدأ التنزيل الآن'); // سطر تصحيح
  } catch (err) {
    console.error('❌ فشل التصدير (ExcelJS):', err); // سطر خطأ
    alert('فشل التصدير: ' + err.message);
    throw err;
  } finally {
    await hideFileOverlay(1200); // إبقاء الأوفرلاي ظاهر ≥ 1.2 ثانية
  }
}


// ================================
// تتبع + تصحيح مسار القالب
// ================================

// سطر الخطأ (يظهر دائمًا عند بدء التصدير/الفحص)
// سطر خطأ لتتبّع المصدر الأصلي للمسار
function logTemplateDebug(rawUrl) {
  console.error('❌ TEMPLATE DEBUG:', {
    rawTpl: rawUrl,
    location: (typeof window !== 'undefined' ? window.location.href : '(no-window)')
  });
}


// ================================
// (نفس الدالة السابقة - لا تغيير)
// ================================
function resolveTemplateUrl() {
  let url = (CONFIG?.EXPORT?.TEMPLATE_URL || 'templates/skill_record.xlsx');

  if (url.startsWith('/')) {
    console.warn('⚠️ تم رصد مسار مطلق — تصحيحه لنسبي');
    url = url.substring(1);
  }

  url = url.replace(/\\/g, '/');
  const stamp = Date.now();
  url += (url.includes('?') ? '&' : '?') + 'v=' + stamp;

  if (typeof window !== 'undefined') {
    if (window.location.protocol === 'file:') {
      console.error('❌ لا يمكن استخدام fetch مع file://');
      alert('⚠️ يرجى تشغيل المشروع عبر HTTP Server');
      return null;
    }

    const abs = new URL(url, window.location.href);
    console.log('✅ مسار القالب النهائي:', abs.href);
    return abs.href;
  }
  
  return url;
}



// دالة فحص سريعة: جرّبها من الكونسول debugTemplateAccess()
async function debugTemplateAccess() {
  try {
    const tplUrl = resolveTemplateUrl();
    logTemplateDebug(CONFIG?.EXPORT?.TEMPLATE_URL);

    console.warn('🔎 فحص الوصول إلى القالب:', tplUrl);
    const resp = await fetch(tplUrl, { method: 'GET', cache: 'no-cache' });
    console.warn('🔎 استجابة الخادم:', resp.status, resp.statusText);

    if (!resp.ok) {
      console.error('❌ فشل التحميل من المسار', { status: resp.status, url: tplUrl });
      console.error('💡 تحقق من: (1) المسار نسبي بدون /، (2) الملف فعليًا تحت templates بجانب index.html، (3) تفتح عبر http وليس file://');
      return;
    }

    const blob = await resp.blob();
    console.warn('🔎 نوع المحتوى:', blob.type || '(غير مذكور)');
    console.warn('🔎 حجم الملف (Byte):', blob.size);
    if (blob.type && !/sheet|excel|octet-stream/i.test(blob.type)) {
      console.error('⚠️ نوع غير متوقع لملف القالب:', blob.type, '— قد لا يقدَّم كـ xlsx بشكل صحيح.');
    } else {
      console.log('✅ ملف القالب قابل للوصول ويبدو صحيحًا.');
    }
  } catch (err) {
    console.error('❌ خطأ شبكة أثناء فحص القالب:', err);
  }
}




async function exportDepartmentExcel() {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getDepartmentExcelData&department=${encodeURIComponent(currentUser.department)}`;
    const res = await fetch(url);
    const data = await res.json();
    
    if (data?.success) {
      // ✅ استخدم نفس دالة التصدير للجميع
      for (const trainee of (data.trainees || [])) {
        // ضَمّن القسم مهما كان اسم العمود الراجع من GAS
        trainee.major = trainee.major || trainee.department || trainee.dept || currentUser?.department || '';
        await exportTraineeExcelUsingTemplate(trainee);
        await new Promise(resolve => setTimeout(resolve, 500));
      }

      
      alert(`✅ تم تصدير ${data.trainees.length} ملف بنجاح!`);
    }
  } catch (e) {
    console.error('خطأ في تصدير Excel:', e);
    alert('حدث خطأ في تصدير الملف');
  } finally {
    spinner?.classList.add('hidden');
  }
}

// ✅ بديل آمن يحافظ على الألوان والدمجات – يستدعي نفس مسار التصدير بالقالب
async function generateExcelReport(data, type) {
  try {
    // سطر تصحيح: استخدم نفس دالة القالب لضمان الحفاظ على التنسيق
    await exportTraineeExcelUsingTemplate({
      name: data?.name || '',
      studentId: data?.studentId || '',
      major: data?.major || '',
      semester: data?.semester || '',
      year: data?.year || '',
      headOfDepartment: data?.headOfDepartment || '',
      deanOfStudents: data?.deanOfStudents || '',
      dean: data?.dean || '',
      // توحيد شكل المهارات: مصفوفة [{name,hours,status}]
      skills: Array.isArray(data?.skills)
        ? data.skills
        : Object.entries(data?.skills || {}).map(([k,v]) => ({ name:k, hours:Number(v||0), status:'معتمد' }))
    });
  } catch (e) {
    console.error('❌ generateExcelReport فشل في استدعاء القالب:', e); // ← سطر الخطأ
    alert('حدث خطأ أثناء إنشاء التقرير من القالب.');
  }
}


// ---- جلسة / تنقّل ----
function logout() {
  localStorage.removeItem('currentUser');
  currentUser = null;
  document.querySelectorAll('.dashboard')?.forEach(el => el.parentElement.classList.add('hidden'));
  document.getElementById('loginPage')?.classList.remove('hidden');
  const u = document.getElementById('username'); if (u) u.value = '';
}

function openAddWorkshopModal() {
  document.getElementById('addWorkshopModal')?.classList.add('active');
}
function closeModal() {
  document.getElementById('addWorkshopModal')?.classList.remove('active');
}

// ---- روابط عرض بسيطة ----
async function viewWorkshopDetailsBasic(workshopId) {
  const modal = document.getElementById('workshopDetailsModal');
  const spinner = document.getElementById('loadingSpinner');
  const titleEl = document.getElementById('wDetName');
  const dateEl  = document.getElementById('wDetDate');
  const locEl   = document.getElementById('wDetLocation');
  const capEl   = document.getElementById('wDetCapacity');
  const statEl  = document.getElementById('wDetStatus');
  const listEl  = document.getElementById('wDetParticipants');

  // سطر خطأ:
  if (!modal || !titleEl || !dateEl || !locEl || !capEl || !statEl || !listEl) {
    console.error('❌ عناصر مودال تفاصيل الورشة غير موجودة في الصفحة.');
    alert('تعذر فتح تفاصيل الورشة: عناصر المودال غير موجودة.');
    return;
  }

  const ws = (window.trainerWorkshops || []).find(w => String(w.id) === String(workshopId));
  // سطر خطأ:
  if (!ws) {
    console.error('❌ لم يتم العثور على الورشة محليًا:', workshopId);
    // سطر تصحيح: جلب طارئ للبيانات ثم إعادة المحاولة
    try {
      await loadTrainerData();
      const ws2 = (window.trainerWorkshops || []).find(w => String(w.id) === String(workshopId));
      if (!ws2) return alert('تعذر العثور على الورشة.');
      return viewWorkshopDetails(workshopId); // إعادة المحاولة بعد التصحيح
    } catch (e) {
      return alert('تعذر تحديث البيانات لعرض التفاصيل.');
    }
  }

  // املأ الهيدر وافتح المودال
  titleEl.textContent = ws.name;
  dateEl.textContent  = ws.date || '-';
  locEl.textContent   = ws.location || '-';
  capEl.textContent   = `${ws.registered || 0}/${ws.capacity || 0}`;
  statEl.textContent  = ws.status || '-';
  statEl.className    = 'workshop-badge ' + ((ws.status === 'نشط' || ws.status === 'متاح') ? 'badge-available' : 'badge-completed');
  modal.classList.add('active');

  // قائمة أولية من المعلقين
  let participants = (window.pendingAttendanceCache || [])
    .filter(a => String(a.workshopName) === String(ws.name))
    .map(a => ({ id:a.id, traineeId:a.traineeId, traineeName:a.traineeName, status:'معلق' }));

  // محاولة جلب تفاصيل كاملة من GAS
  spinner?.classList.remove('hidden');
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getWorkshopDetails&workshopId=${encodeURIComponent(workshopId)}`;
    const res = await fetch(url);
    const data = await res.json();

    if (data?.success && Array.isArray(data.participants)) {
      participants = data.participants.map(p => ({
        id: p.attendanceId,
        traineeId: p.traineeId,
        traineeName: p.traineeName,
        status: p.status
      }));
    } else if (data && data.message === 'Invalid action') {
      console.log('ℹ️ لا يوجد getWorkshopDetails في GAS — تم استخدام البيانات المحلية.');
    }
  } catch (e) {
    console.error('خطأ عند جلب تفاصيل الورشة من GAS:', e); // سطر خطأ
  } finally {
    spinner?.classList.add('hidden');
  }

  renderWorkshopParticipants(document.getElementById('wDetParticipants'), participants);
}



// ✅ توليد شهادة PNG محليًا للمهارة "المعتمدة"
function downloadCertificate(skillId) {
  try {
    const skill = (window.traineeSkillsCache || [])[Number(skillId)] // لأنك استخدمت i كـ id في الجدول
                 || (window.traineeSkillsCache || []).find(s => String(s.id) === String(skillId));

    if (!skill) {
      console.error('❌ لم يتم العثور على السجل المطلوب للشهادة:', skillId); // سطر خطأ
      return alert('تعذر العثور على بيانات الشهادة.');
    }
    if (skill.status !== 'معتمد') {
      return alert('الشهادة متاحة فقط للسجلات المعتمدة.');
    }


    // خصائص الشهادة
    const traineeName = currentUser?.name || 'المتدرب';
    const skillName   = skill.name || 'مهارة';
    const hours       = skill.hours || 0;
    const dateTxt     = skill.date || '';

    // Canvas
    const W = 1280, H = 900;
    const c = document.createElement('canvas');
    c.width = W; c.height = H;
    const ctx = c.getContext('2d');

    // خلفية بسيطة
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0,0,W,H);

    // إطار
    ctx.strokeStyle = '#186F65';
    ctx.lineWidth = 12;
    ctx.strokeRect(30,30,W-60,H-60);

    // ترويسة
    ctx.fillStyle = '#186F65';
    ctx.font = 'bold 42px "Tahoma"';
    ctx.fillText('المؤسسة العامة للتدريب التقني والمهني', 80, 110);

    ctx.fillStyle = '#333';
    ctx.font = 'bold 56px "Tahoma"';
    ctx.fillText('شهادة إتمام مهارة', 80, 190);

    // نص الشهادة
    ctx.font = '28px "Tahoma"';
    ctx.fillText(`يُمنح ${traineeName} هذه الشهادة لإتمامه مهارة:`, 80, 270);

    ctx.font = 'bold 36px "Tahoma"';
    ctx.fillText(skillName, 80, 320);

    ctx.font = '28px "Tahoma"';
    ctx.fillText(`عدد الساعات: ${hours}`, 80, 380);
    ctx.fillText(`التاريخ: ${dateTxt}`, 80, 430);

    // توقيعات/جهات
    ctx.font = '24px "Tahoma"';
    ctx.fillText('رئيس القسم:', 80, H-160);
    ctx.fillText('وكيل شؤون المتدربين:', 80, H-120);
    ctx.fillText('العميد:', 80, H-80);

    // تحميل
    const link = document.createElement('a');
    link.download = `شهادة_${traineeName}_${skillName}.png`;
    link.href = c.toDataURL('image/png');
    link.click();
  } catch (e) {
    console.error('خطأ أثناء توليد الشهادة:', e); // ← سطر الخطأ
    alert('حدث خطأ أثناء توليد الشهادة.');
  }
}

function viewDepartmentDetails(departmentName){ alert('سيتم فتح تفاصيل قسم: ' + departmentName); }

// ---- عند تحميل الصفحة: استعادة الجلسة مع التطبيع ----
window.onload = function () {
  const saved = localStorage.getItem('currentUser');
  if (saved) {
    try {
      currentUser = JSON.parse(saved);
      currentUser.userType = normalizeUserType(currentUser.userType);
      localStorage.setItem('currentUser', JSON.stringify(currentUser));
      redirectToDashboard(currentUser.userType);
    } catch {
      localStorage.removeItem('currentUser');
    }
  }
};


// عدّل loadTrainerData ليحفظ البيانات في الذاكرة:
async function loadTrainerData() {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getTrainerData&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    const data = await res.json();
    if (data?.success) {
      window.trainerWorkshops = data.workshops || [];
      window.pendingAttendanceCache = data.pendingAttendance || [];
      renderTrainerStats(data.stats);
      renderTrainerWorkshops(data.workshops);
      renderPendingAttendance(data.pendingAttendance);
    } else {
      showError(data?.message || 'تعذر جلب بيانات المدرب');
    }
  } catch (e) {
    console.error('خطأ في جلب البيانات:', e);
  } finally {
    spinner?.classList.add('hidden');
  }
}

// ✅ دالة إظهار تفاصيل الورشة داخل مودال
async function viewWorkshopDetails(workshopId) {
  const modal = document.getElementById('workshopDetailsModal');
  const spinner = document.getElementById('loadingSpinner');
  const titleEl = document.getElementById('wDetName');
  const dateEl  = document.getElementById('wDetDate');
  const locEl   = document.getElementById('wDetLocation');
  const capEl   = document.getElementById('wDetCapacity');
  const statEl  = document.getElementById('wDetStatus');
  const listEl  = document.getElementById('wDetParticipants');

  // فحص وجود عناصر المودال
  if (!modal || !titleEl || !dateEl || !locEl || !capEl || !statEl || !listEl) {
    console.error('❌ عناصر مودال تفاصيل الورشة غير موجودة في الصفحة.'); // ← سطر الخطأ المطلوب
    alert('تعذر فتح تفاصيل الورشة: عناصر المودال غير موجودة.');
    return;
  }

  // ابحث عن الورشة محليًا
  const ws = (window.trainerWorkshops || []).find(w => String(w.id) === String(workshopId));
  if (!ws) {
    console.error('❌ لم يتم العثور على الورشة محليًا:', workshopId); // ← سطر الخطأ
    alert('تعذر العثور على الورشة محليًا.');
    return;
  }

  // املأ الهيدر
  titleEl.textContent = ws.name;
  dateEl.textContent  = ws.date || '-';
  locEl.textContent   = ws.location || '-';
  capEl.textContent   = `${ws.registered || 0}/${ws.capacity || 0}`;
  statEl.textContent  = ws.status || '-';
  statEl.className    = 'workshop-badge ' + ((ws.status === 'نشط' || ws.status === 'متاح') ? 'badge-available' : 'badge-completed');

  // افتح المودال فورًا (سلوك سريع للمستخدم)
  modal.classList.add('active');

  // حضّر قائمة مبدئية من "المعلقين" من الكاش المحلي
  let participants = (window.pendingAttendanceCache || [])
    .filter(a => String(a.workshopName) === String(ws.name))
    .map(a => ({ id:a.id, traineeId:a.traineeId, traineeName:a.traineeName, status:'معلق' }));

  // جرّب إن كان لديك أكشن Backend يُرجع التفاصيل الكاملة (معتمد + معلق)
  spinner?.classList.remove('hidden');
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getWorkshopDetails&workshopId=${encodeURIComponent(workshopId)}`;
    const res = await fetch(url);
    const data = await res.json();

    if (data?.success && Array.isArray(data.participants)) {
      // لو وجدنا الأكشن في GAS، استخدمه بدل المبدئي
      participants = data.participants.map(p => ({
        id: p.attendanceId,
        traineeId: p.traineeId,
        traineeName: p.traineeName,
        status: p.status // معتمد / معلق / ملغي...
      }));
    } else if (data && data.message === 'Invalid action') {
      // لا يوجد أكشن في GAS؛ لا مشكلة — أبقِ على المبدئي
      console.log('ℹ️ لا يوجد getWorkshopDetails في GAS — تم استخدام البيانات المحلية.');
    }
  } catch (e) {
    console.error('خطأ عند جلب تفاصيل الورشة من GAS:', e); // ← سطر الخطأ
  } finally {
    spinner?.classList.add('hidden');
  }

  // ارسم القائمة
  renderWorkshopParticipants(listEl, participants);
}

// رسم جدول/قائمة المشاركين + أزرار اعتماد سريعة
function renderWorkshopParticipants(containerEl, participants) {
  if (!participants || !participants.length) {
    containerEl.innerHTML = `<p style="text-align:center;color:var(--tvtc-text-muted);">لا يوجد مسجلون حتى الآن.</p>`;
    return;
  }

  // دوال مساعدة صغيرة لعرض الشارة واختيار اللائحة
  const badgeFor = (statusAr) => {
    if (statusAr === 'معتمد') return 'workshop-badge badge-completed';
    if (statusAr === 'معلق')  return 'workshop-badge badge-pending';
    if (statusAr === 'مرفوض' || statusAr === 'ملغي' || statusAr === 'لم يحضر') return 'workshop-badge badge-error';
    return 'workshop-badge badge-available';
  };

  const selectHtml = (id, statusAr) => {
    // نطابق المعاني:
    const isApproved = statusAr === 'معتمد';
    const isPending  = statusAr === 'معلق';
    const isNoShow   = (statusAr === 'مرفوض' || statusAr === 'ملغي' || statusAr === 'لم يحضر');

    return `
      <select class="input input-small" onchange="setAttendanceStatusUI('${id}', this)">
        <option value="pending"  ${isPending  ? 'selected' : ''}>معلق</option>
        <option value="approved" ${isApproved ? 'selected' : ''}>معتمد (حضور)</option>
        <option value="noshow"   ${isNoShow   ? 'selected' : ''}>لم يحضر</option>
      </select>
    `;
  };

  const rows = participants.map(p => `
    <tr>
      <td>${p.traineeName}</td>
      <td>${p.traineeId}</td>
      <td>
        <span class="${badgeFor(p.status)}">${p.status}</span>
      </td>
      <td>
        ${selectHtml(p.id, p.status)}
      </td>
    </tr>
  `).join('');

  containerEl.innerHTML = `
    <table>
      <thead>
        <tr><th>اسم المتدرب</th><th>رقم المتدرب</th><th>الحالة</th><th>إجراء</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

window.exportTraineeExcel = exportTraineeExcel;

// زر إغلاق مودال تفاصيل الورشة
function closeWorkshopDetailsModal() {
  document.getElementById('workshopDetailsModal')?.classList.remove('active');
}
