// ============================================
// السكريبت الرئيسي - app.js (نسخة مصححة)
// ============================================
// إصلاح مشكلة pdfmake vFS
if (typeof pdfMake !== 'undefined') {
    if (!pdfMake.vFS && pdfMake.vfs) {
        pdfMake.vFS = pdfMake.vfs;
    }
    
    // تهيئة الخطوط الافتراضية إذا لم تكن موجودة
    if (!pdfMake.fonts) {
        pdfMake.fonts = {
            Roboto: {
                normal: 'Roboto-Regular.ttf',
                bold: 'Roboto-Medium.ttf',
                italics: 'Roboto-Italic.ttf',
                bolditalics: 'Roboto-MediumItalic.ttf'
            }
        };
    }
}

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
      currentUser = { 
        ...data.user, 
        userType: normalizeUserType(data.user.userType) 
      };
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
  upgradeIcons(document.getElementById('traineeStats'));

}


function renderAvailableWorkshops(workshops) {
  const html = (workshops || []).map(w => `
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
          : `<button class="btn btn-outline btn-small" disabled>غير متاح</button>`}
        <!-- ⚠️ مُزال: لا زر تفاصيل للمتدرب -->
      </div>
    </div>
  `).join('');
  document.getElementById('availableWorkshops').innerHTML = html;
  upgradeIcons(document.getElementById('availableWorkshops'));
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
                ? `<button class="btn btn-accent btn-small" onclick="downloadCertificatePDF('${s.id}')">تحميل</button>`
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
  upgradeIcons(document.getElementById('trainerStats'));

}

function renderTrainerWorkshops(workshops) {
  const html = `
    <div class="table-responsive">
      <table>
        <thead>
          <tr>
            <th>اسم الورشة</th>
            <th>التاريخ</th>
            <th>المكان</th>
            <th>المسجلين</th>
            <th>الحالة</th>
            <th>الإجراءات</th>
          </tr>
        </thead>
        <tbody>
          ${(workshops||[]).map(w => {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const wDate = parseDateFlexible_(w.date);
            const isPast = wDate && wDate < today;
            const statusBadge = isPast ? 'badge-error' : 
                               (w.status === 'نشط' || w.status === 'متاح') ? 'badge-available' : 
                               'badge-completed';
            
            return `
              <tr>
                <td>
                  <strong>${w.name}</strong>
                  ${isPast ? '<span style="color:var(--error);font-size:12px;margin-right:8px;">⏱️ منتهية</span>' : ''}
                </td>
                <td>${w.date}</td>
                <td>${w.location}</td>
                <td>
                  <span style="font-weight:600;">
                    ${w.registered}/${w.capacity}
                  </span>
                </td>
                <td>
                  <span class="workshop-badge ${statusBadge}">
                    ${w.status}
                  </span>
                </td>
                <td>
                  <button class="btn btn-primary btn-small" 
                          onclick="viewWorkshopDetails('${w.id}')">
                    📋 التفاصيل
                  </button>
                </td>
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
  document.getElementById('trainerWorkshops').innerHTML = html;
}

// ---- تحميل/جاهزية الخطوط للـ HTML2Canvas ----
async function ensureArabicWebFontsReady() {
  try {
    if (document?.fonts?.ready) {
      await document.fonts.ready; // ينتظر كل الويب فونتس
    } else {
      // متصفحات قديمة: مهلة بسيطة كـ fallback
      await new Promise(r => setTimeout(r, 400));
    }
  } catch { /* تجاهل */ }
}


// ✅ عرض الورش النشطة مع متدربيها المعلقين فقط
function renderPendingAttendance(attendance) {
  const container = document.getElementById('pendingAttendance');
  if (!container) return;
  
  if (!attendance || !attendance.length) {
    container.innerHTML = `
      <div style="text-align:center; padding: 40px; color:var(--tvtc-text-muted);">
        <div style="font-size: 48px; margin-bottom: 16px;">✅</div>
        <h3 style="margin: 0 0 8px 0;">رائع! لا توجد سجلات معلقة</h3>
        <p style="margin: 0;">تم اعتماد جميع الحضور</p>
      </div>
    `;
    return;
  }

  // ✅ تجميع المتدربين حسب الورشة
  const byWorkshop = {};
  attendance.forEach(a => {
    const wsName = a.workshopName || 'ورشة غير محددة';
    if (!byWorkshop[wsName]) {
      byWorkshop[wsName] = [];
    }
    byWorkshop[wsName].push(a);
  });

  // ✅ بناء HTML منظم
  let html = '<div class="workshops-approval-container">';
  
  Object.keys(byWorkshop).forEach((workshopName, idx) => {
    const trainees = byWorkshop[workshopName];
    const workshopId = `workshop-${idx}`;
    
    html += `
      <div class="workshop-approval-card">
        <!-- Header -->
        <div class="workshop-header">
          <div>
            <h3>📋 ${workshopName}</h3>
            <p style="margin: 4px 0 0 0; opacity: 0.9; font-size: 14px;">
              ${trainees.length} متدرب في انتظار الاعتماد
            </p>
          </div>
        </div>
        
        <!-- شريط الأدوات -->
        <div class="workshop-actions-header">
          <label class="checkbox-label">
            <input type="checkbox" 
                   id="selectAll-${workshopId}"
                   class="select-all-checkbox" 
                   onchange="toggleAllInWorkshop('${workshopId}', this)">
            <span>تحديد الكل (${trainees.length})</span>
          </label>
          
          <div class="action-buttons">
            <button class="btn btn-primary btn-small" 
                    onclick="approveWorkshopAttendance('${workshopId}', 'approve')">
              ✅ اعتماد الحضور
            </button>
            <button class="btn btn-outline btn-small" 
                    onclick="approveWorkshopAttendance('${workshopId}', 'noshow')"
                    style="border-color: var(--error); color: var(--error);">
              ❌ لم يحضر
            </button>
          </div>
        </div>

        <!-- الجدول -->
        <div class="table-wrapper">
          <table class="trainees-table" id="${workshopId}">
            <thead>
              <tr>
                <th style="width: 50px; text-align: center;">تحديد</th>
                <th>اسم المتدرب</th>
                <th>رقم التدريب</th>
                <th>وقت التسجيل</th>
              </tr>
            </thead>
            <tbody>`;
    
    trainees.forEach(t => {
      html += `
        <tr class="trainee-row">
          <td style="text-align: center;">
            <input type="checkbox" 
                   class="trainee-checkbox" 
                   value="${t.id}"
                   onchange="updateSelectAllState('${workshopId}')">
          </td>
          <td><strong>${t.traineeName}</strong></td>
          <td>${t.traineeId}</td>
          <td>${t.registrationTime}</td>
        </tr>`;
    });
    
    html += `
            </tbody>
          </table>
        </div>
        
        <!-- عداد المحدد -->
        <div class="workshop-footer" id="footer-${workshopId}">
          <span class="selected-count">لم يتم تحديد أي متدرب</span>
        </div>
      </div>`;
  });
  
  html += '</div>';
  container.innerHTML = html;
}


// ✅ تحديث حالة checkbox "تحديد الكل" والعداد
function updateSelectAllState(workshopId) {
  const table = document.getElementById(workshopId);
  const checkboxes = table.querySelectorAll('.trainee-checkbox');
  const checkedBoxes = table.querySelectorAll('.trainee-checkbox:checked');
  const selectAllCheckbox = document.getElementById(`selectAll-${workshopId}`);
  const footer = document.getElementById(`footer-${workshopId}`);
  
  // تحديث checkbox الرئيسي
  if (selectAllCheckbox) {
    selectAllCheckbox.checked = checkedBoxes.length === checkboxes.length && checkboxes.length > 0;
    selectAllCheckbox.indeterminate = checkedBoxes.length > 0 && checkedBoxes.length < checkboxes.length;
  }
  
  // تحديث العداد
  if (footer) {
    const count = checkedBoxes.length;
    if (count === 0) {
      footer.innerHTML = '<span class="selected-count">👆 حدد المتدربين الذين حضروا</span>';
    } else {
      footer.innerHTML = `
        <span class="selected-count selected">
          ✓ تم تحديد <strong>${count}</strong> من ${checkboxes.length} متدرب
        </span>
      `;
    }
  }
}


// ✅ النسخة الموحّدة لاعتماد حضور المتدربين في ورشة محددة
async function approveWorkshopAttendance(workshopId, action) {
  // تحقق من وجود الجدول
  const table = document.getElementById(workshopId);
  if (!table) {
    alert('تعذر الوصول إلى قائمة المتدربين لهذه الورشة.');
    return;
  }

  // اجمع المحددين
  const checked = table.querySelectorAll('.trainee-checkbox:checked');
  if (checked.length === 0) {
    alert('⚠️ يرجى تحديد المتدربين الذين حضروا باستخدام ☑️');
    return;
  }

  // نصوص التأكيد والأيقونة
  const isApprove = action === 'approve';
  const actionText = isApprove ? 'اعتماد حضور' : 'تعليم "لم يحضر" لـ';
  const icon = isApprove ? '✅' : '❌';

  const ok = window.confirm(`${icon} هل أنت متأكد من ${actionText} ${checked.length} متدرب؟`);
  if (!ok) return;

  // إظهار حالة الانشغال
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  showFileOverlay('⚡ جاري المعالجة...', `معالجة ${checked.length} سجل`);

  try {
    let successCount = 0;
    const statusValue = isApprove ? 'approved' : 'noshow';

    // طبّق الحالة لكل عنصر محدد
    for (const cb of checked) {
      const attendanceId = cb.value;
      const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({
          action: 'setAttendanceStatus',
          attendanceId,
          trainerId: currentUser.id,
          status: statusValue
        })
      });
      const data = await res.json();
      if (data?.success) successCount++;
    }

    // حدّث الـ overlay برسالة نجاح
    updateFileOverlay('✅ تمت المعالجة بنجاح', `تم ${actionText} ${successCount} سجل`);

    // إعادة تحميل بيانات المدرب وتحديث الواجهة
    setTimeout(async () => {
      hideFileOverlay(0);
      await loadTrainerData();
      alert(`✅ تم ${actionText} ${successCount} من ${checked.length} سجل بنجاح!`);
    }, 1500);
  } catch (err) {
    console.error('خطأ في الاعتماد:', err);
    updateFileOverlay('❌ فشلت العملية', err.message || 'تعذر إكمال العملية', true);
    setTimeout(() => hideFileOverlay(0), 3000);
  } finally {
    spinner?.classList.add('hidden');
  }
}



// ✅ تحديد/إلغاء تحديد الكل
function toggleAllInWorkshop(workshopId, checkbox) {
  const table = document.getElementById(workshopId);
  const checkboxes = table.querySelectorAll('.trainee-checkbox');
  checkboxes.forEach(cb => cb.checked = checkbox.checked);
  updateSelectAllState(workshopId);
}



function toggleWorkshopAttendance(workshopId, checked) {
  const table = document.getElementById(workshopId);
  const checkboxes = table.querySelectorAll('.trainee-checkbox');
  checkboxes.forEach(cb => cb.checked = checked);
  
  // تحديث checkbox الرئيسي
  const mainCheckbox = table.querySelector('.workshop-select-all');
  if (mainCheckbox) mainCheckbox.checked = checked;
}


// دوال مساعدة للـ checkboxes
function toggleAllAttendance(checkbox) {
  const checkboxes = document.querySelectorAll('.attendance-checkbox');
  checkboxes.forEach(cb => cb.checked = checkbox.checked);
}

async function bulkApproveSelected(action) {
  const checkboxes = document.querySelectorAll('.attendance-checkbox:checked');
  
  if (checkboxes.length === 0) {
    alert('⚠️ يرجى تحديد سجل واحد على الأقل');
    return;
  }
  
  const actionText = action === 'approve' ? 'اعتماد' : 'تعليم كـ "لم يحضر" لـ';
  const confirm = window.confirm(
    `هل أنت متأكد من ${actionText} ${checkboxes.length} سجل محدد؟`
  );
  
  if (!confirm) return;
  
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  
  try {
    showFileOverlay('⚡ جاري المعالجة...', `معالجة ${checkboxes.length} سجل`);
    
    let successCount = 0;
    const statusValue = action === 'approve' ? 'approved' : 'noshow';
    
    for (const checkbox of checkboxes) {
      const attendanceId = checkbox.value;
      
      const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({
          action: 'setAttendanceStatus',
          attendanceId,
          trainerId: currentUser.id,
          status: statusValue
        })
      });
      
      const data = await res.json();
      if (data?.success) successCount++;
    }
    
    updateFileOverlay('✅ تمت المعالجة', `تم ${actionText} ${successCount} سجل بنجاح`);
    
    setTimeout(async () => {
      hideFileOverlay(0);
      await loadTrainerData();
      alert(`✅ تم ${actionText} ${successCount} من ${checkboxes.length} سجل بنجاح!`);
    }, 2000);
    
  } catch (err) {
    console.error('خطأ في الاعتماد:', err);
    updateFileOverlay('❌ فشلت العملية', err.message, true);
    setTimeout(() => hideFileOverlay(0), 3000);
  } finally {
    spinner?.classList.add('hidden');
  }
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

// ================================
// أيقونات احترافية: ترقية تلقائية من الإيموجي إلى SVG
// ================================
function upgradeIcons(root=document){
  // خريطة تحويل الإيموجي -> id الرمز في الـ sprite
  const map = {
    '⏱️':'i-clock',
    '📅':'i-calendar',
    '📍':'i-map',
    '👥':'i-users',
    '🏆':'i-trophy',
    '📚':'i-trophy' // بديل جميل للعنوان
  };

  // استبدال كل span أول داخل .workshop-detail (الذي يحوي الإيموجي)
  root.querySelectorAll('.workshop-detail span:first-child').forEach(el=>{
    const t = (el.textContent||'').trim();
    const id = map[t];
    if (!id) return;
    const svg = document.createElementNS('http://www.w3.org/2000/svg','svg');
    svg.classList.add('i');
    const use = document.createElementNS('http://www.w3.org/2000/svg','use');
    use.setAttributeNS('http://www.w3.org/1999/xlink','href','#'+id);
    svg.appendChild(use);
    // لف الأيقونة مع النص الحالي المجاور (إن وُجد)
    el.replaceWith(svg);
  });

  // ترقية أي رموز في عناوين الأقسام/الإحصاءات
  root.querySelectorAll('.section-title, .stat-icon').forEach(el=>{
    const txt = (el.textContent||'').trim();
    const key = Object.keys(map).find(k=>txt.startsWith(k));
    if (!key) return;
    // إبقِ النص بعد مسح الإيموجي الأولى
    el.textContent = txt.replace(key,'').trim();
    // أضف SVG قبل النص
    const svg = document.createElementNS('http://www.w3.org/2000/svg','svg');
    svg.classList.add('i');
    const use = document.createElementNS('http://www.w3.org/2000/svg','use');
    use.setAttributeNS('http://www.w3.org/1999/xlink','href','#'+map[key]);
    svg.appendChild(use);
    const wrap = document.createElement('span');
    wrap.className = 'with-icon';
    wrap.appendChild(svg);
    wrap.append(' ' + el.textContent);
    el.replaceWith(wrap);
  });
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

// في دالة loadHeadData، أضف:
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
      
      // جلب الشهادات المعلقة
      await loadPendingExternalCertsForHead();
    }
  } catch (e) {
    console.error('خطأ في جلب البيانات:', e);
  } finally {
    spinner?.classList.add('hidden');
  }
}

async function loadPendingExternalCertsForHead() {
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=listExternalCertificates&department=${encodeURIComponent(currentUser.department)}&status=بانتظار الاعتماد`;
    const res = await fetch(url);
    const data = await res.json();
    
    if (data?.success) {
      renderPendingExternalCerts(data.certificates || []);
    }
  } catch (e) {
    console.error('خطأ في جلب الشهادات المعلقة:', e);
  }
}

function renderPendingExternalCerts(certs) {
  const container = document.getElementById('pendingExternalCerts');
  if (!container) return;
  
  if (!certs.length) {
    container.innerHTML = '<p style="text-align:center;color:var(--tvtc-text-muted);">لا توجد شهادات بانتظار الاعتماد.</p>';
    return;
  }
  
  const rows = certs.map(c => `
    <tr>
      <td><strong>${c.traineeName}</strong></td>
      <td>${c.courseName}</td>
      <td>${c.hours} ساعة</td>
      <td>
        ${c.fileUrl ? `<a href="${c.fileUrl}" target="_blank" class="btn btn-outline btn-small">📎 عرض</a>` : '-'}
      </td>
      <td>
        <button class="btn btn-primary btn-small" onclick="approveExternalCert('${c.id}', true)">✅ اعتماد</button>
        <button class="btn btn-outline btn-small" onclick="approveExternalCert('${c.id}', false)" style="margin-right:5px;">❌ رفض</button>
      </td>
    </tr>
  `).join('');
  
  container.innerHTML = `
    <table>
      <thead>
        <tr><th>المتدرب</th><th>اسم الدورة</th><th>الساعات</th><th>الملف</th><th>الإجراء</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

async function approveExternalCert(certId, approve) {
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  
  try {
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'approveExternalCertificate',
        certId,
        approverId: currentUser.id,
        approve
      })
    });
    
    const data = await res.json();
    
    if (data?.success) {
      alert(approve ? '✅ تم اعتماد الشهادة وإضافتها للمهارات' : '❌ تم رفض الشهادة');
      await loadHeadData(); // إعادة تحميل
    } else {
      alert('خطأ: ' + (data?.message || 'فشل العملية'));
    }
  } catch (e) {
    console.error('خطأ في اعتماد الشهادة:', e);
    alert('حدث خطأ في العملية');
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
// ============================================
// تصدير سجل المهارات المحسّن مع مزامنة آمنة
// ============================================

async function exportTraineeExcel() {
  try {
    if (!currentUser || !currentUser.id) {
      showError('يجب تسجيل الدخول أولاً');
      return;
    }

    showFileOverlay('📋 جاري الاتصال بالخادم...', 'نحضّر بياناتك من السحابة');

    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=exportTraineeXlsx&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    
    if (!res.ok) {
      throw new Error(`فشل الاتصال بالخادم: ${res.status} ${res.statusText}`);
    }

    const data = await res.json();
    
    if (!data.success) {
      throw new Error(data.message || 'تعذر جلب البيانات من الخادم');
    }

    if (!data.skills || data.skills.length === 0) {
      updateFileOverlay('⚠️ تنبيه', 'لا توجد مهارات معتمدة للتصدير', true);
      setTimeout(() => hideFileOverlay(0), 3000);
      return;
    }

    updateFileOverlay('✅ تم استلام البيانات', 'جاري تحميل القالب...');

    const tplUrl = resolveTemplateUrl();
    if (!tplUrl) {
      throw new Error('تعذر تحديد مسار قالب Excel');
    }

    const tplRes = await fetch(tplUrl, { cache: 'no-cache' });
    if (!tplRes.ok) {
      throw new Error(`فشل تحميل القالب: ${tplRes.status}`);
    }

    const tplBuf = await tplRes.arrayBuffer();
    updateFileOverlay('📊 جاري معالجة البيانات...', 'كتابة المهارات والتوقيعات');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(tplBuf);

    const ws = wb.getWorksheet(CONFIG.EXPORT.SHEET_NAME);
    if (!ws) {
      throw new Error('ورقة القالب غير موجودة داخل الملف');
    }

      // 🔒 قفل جميع الخلايا (افتراضياً)
      ws.eachRow(row => {
        row.eachCell(cell => {
          cell.protection = { locked: true };
        });
      });

      // ✅ فعّل الحماية (انتبه: Promise)
      await ws.protect('TVTC2025', {
        selectLockedCells: true,
        selectUnlockedCells: true,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertRows: false,
        insertColumns: false,
        deleteRows: false,
        deleteColumns: false,
        sort: false,
        autoFilter: false,
        pivotTables: false
      });


    const traineeData = data.traineeData || {};
    const skillsArr = data.skills || [];

    // كتابة معلومات المتدرب
    ws.getCell(CONFIG.EXPORT.NAME_CELL).value  = traineeData.name || '';
    ws.getCell(CONFIG.EXPORT.ID_CELL).value    = traineeData.studentId || '';
    ws.getCell(CONFIG.EXPORT.MAJOR_CELL).value = traineeData.major || traineeData.department || '';
    ws.getCell(CONFIG.EXPORT.SEM_CELL).value   = traineeData.semester || '';
    ws.getCell(CONFIG.EXPORT.YEAR_CELL).value  = traineeData.year || '';

    // معالجة التوقيعات
    traineeData.department = traineeData.major || traineeData.department || currentUser?.department || '';
    updateFileOverlay('🔐 جاري التحقق من التوقيعات...', 'نبحث عن المسؤولين');
    
    await resolveSignaturesOnlineStrict(traineeData);

    ws.getCell(CONFIG.EXPORT.HOD_CELL).value      = traineeData.headOfDepartment || 'رئيس القسم';
    ws.getCell(CONFIG.EXPORT.DEAN_STD_CELL).value = traineeData.deanOfStudents || 'وكيل شؤون المتدربين';
    ws.getCell(CONFIG.EXPORT.DEAN_CELL).value     = traineeData.dean || 'العميد';

    // ✅ كتابة المهارات متتالية بدون فراغات
    updateFileOverlay('📝 جاري كتابة المهارات...', `معالجة ${skillsArr.length} مهارة معتمدة`);
    
    let totalHours = 0;
    const startRow = 9; // بداية كتابة المهارات
    
    skillsArr.forEach((skill, index) => {
      const currentRow = startRow + index;
      const hours = Number(skill.hours || 0);
      
      if (hours > 0) {
        // كتابة اسم المهارة في B
        ws.getCell(`B${currentRow}`).value = skill.name;
        
        // كتابة الساعات في F
        ws.getCell(`F${currentRow}`).value = hours;
        
        totalHours += hours;
        
        console.log(`  ✓ الصف ${currentRow}: ${skill.name} - ${hours} ساعة`);
      }
    });

    // كتابة الإجمالي في الصف التالي
    // ✅ كتابة المجموع في الخلية الثابتة F30
    ws.getCell(CONFIG.EXPORT.TOTAL_HOURS_CELL || 'F30').value = totalHours;
    
    console.log(`📊 إجمالي الساعات: ${totalHours}`);

    updateFileOverlay('📦 جاري تحضير الملف للتنزيل...', 'لحظات وسيبدأ التحميل');
    
    const outBuf = await wb.xlsx.writeBuffer();
    const safeId = String(traineeData.studentId || currentUser.id || '').replace(/[^\w]/g, '');
    const fileName = `سجل_المهارات_${safeId}_${Date.now()}.xlsx`;

    const blob = new Blob([outBuf], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 100);

    updateFileOverlay('✅ تم إنشاء الملف بنجاح!', 'يبدأ التنزيل الآن...');

  } catch (err) {
    console.error('❌ فشل التصدير:', err);
    updateFileOverlay('⚠️ حدث خطأ', err.message || 'فشل إنشاء ملف Excel', true);
    setTimeout(() => hideFileOverlay(0), 3000);
    return;
  }

  await hideFileOverlay(2000);
}

if (typeof DriveApp !== 'undefined') {
// رفع ملف إلى Google Drive
function saveFileToDrive_(fileName, mimeType, base64Content) {
  try {
    const folder = getDriveFolder_();
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    console.error('خطأ في حفظ الملف:', e);
    return '';
  }
}

function getDriveFolder_() {
  const folderName = 'الشهادات_الخارجية';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

// تحديث submitExternalCertificate لحفظ الملف
function submitExternalCertificate(userId, courseName, hours, fileName, mimeType, fileContent) {
  const usersSheet = getSheet_(SHEETS.USERS);
  if (!usersSheet) return jsonResponse(false, 'ورقة المستخدمين غير موجودة');

  const values = usersSheet.getDataRange().getValues();
  let traineeName = '', dept = '';
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][COL.USERS.ID]) === String(userId)) {
      traineeName = String(values[i][COL.USERS.NAME]||'').trim();
      dept = String(values[i][COL.USERS.DEPT]||'').trim();
      break;
    }
  }
  if (!traineeName) return jsonResponse(false, 'المتدرب غير موجود');

  // حفظ الملف في Drive
  const fileUrl = saveFileToDrive_(fileName, mimeType, fileContent);
  
  const h = ensureExternalCertsSheet_();
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const id = nextId_('CERT');
    h.appendRow([
      id, userId, traineeName, dept,
      String(courseName||'').trim(),
      toNumber_(hours),
      fileUrl,
      CERT_STATUS.PENDING, '', ''
    ]);
    return jsonResponse(true, 'تم رفع الشهادة وبانتظار اعتماد رئيس القسم', { certId: id });
  } finally {
    lock.releaseLock();
  }
}
}
// ============================================
// دوال مساعدة محسّنة للـ Overlay
// ============================================

async function hideFileOverlay(delayMs = 1500) {
  if (delayMs > 0) {
    await new Promise(resolve => setTimeout(resolve, delayMs));
  }
  const ov = document.getElementById('fileDownloadOverlay');
  if (ov) ov.classList.add('hidden');
}

function updateFileOverlay(title, subtitle, isError = false) {
  const ov = document.getElementById('fileDownloadOverlay');
  const titleEl = document.getElementById('fileOverlayTitle');
  const msgEl = document.getElementById('fileOverlayMsg');
  
  if (!ov || !titleEl || !msgEl) return;
  
  titleEl.textContent = title;
  msgEl.textContent = subtitle;
  
  if (isError) {
    ov.classList.add('error-state');
  } else {
    ov.classList.remove('error-state');
  }
}

let __fileOverlayStart = 0;

function showFileOverlay(title = 'جاري التحضير...', subtitle = '') {
  const ov = document.getElementById('fileDownloadOverlay');
  const titleEl = document.getElementById('fileOverlayTitle');
  const msgEl = document.getElementById('fileOverlayMsg');
  
  if (!ov || !titleEl || !msgEl) {
    console.error('❌ عناصر الـ overlay غير موجودة');
    return;
  }

  __fileOverlayStart = Date.now();
  titleEl.textContent = title;
  msgEl.textContent = subtitle;
  
  // إزالة حالة الخطأ
  ov.classList.remove('error-state');
  ov.classList.remove('hidden');
}

// ✅ تصدير PDF محسّن باستخدام html2canvas + jsPDF
async function exportTraineePDF() {
  if (!currentUser || !currentUser.id) {
    showError('يجب تسجيل الدخول أولاً');
    return;
  }

  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');

  try {
    showFileOverlay('📄 جاري إنشاء ملف PDF...', 'قد يستغرق بضع ثوانٍ');

    // 1) جلب بيانات المتدرب والمهارات
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=exportTraineeXlsx&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`فشل الاتصال بالخادم: ${res.status}`);
    const data = await res.json();
    if (!data.success || !Array.isArray(data.skills) || data.skills.length === 0) {
      throw new Error('لا توجد مهارات معتمدة للتصدير');
    }

    updateFileOverlay('✅ تم استلام البيانات', 'تجهيز القالب...');

    // 2) تجهيز بيانات التواقيع
    const traineeData = data.traineeData || {};
    const skills = data.skills || [];
    await resolveSignaturesOnlineStrict(traineeData);

    // 3) بناء الـ HTML الخاص بالطباعة
    const htmlContent = createPDFHTML(traineeData, skills);

    // 4) حقنه في حاوية مخفية
    const pdfContainer = document.getElementById('pdfExportContainer') || createPDFContainer();
    pdfContainer.innerHTML = htmlContent;

    // 5) انتظار تحميل الخطوط
    await ensureArabicWebFontsReady();

    // 6) تحويل HTML إلى صورة
    const canvas = await html2canvas(pdfContainer, {
      scale: 2,
      useCORS: true,
      allowTaint: false,
      backgroundColor: '#ffffff',
      logging: false
    });

    // 7) إنشاء jsPDF + إضافة خط Tajawal على كائن المستند (✅ التصحيح هنا)
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    const tajawalFontBase64 = `
AAEAAAASAQAABAAgR0RFRrRCsIIAAAHgAAAAHEdQT1OyLWdKAAAB4AAAAExHU1VCZ0/fOwAAAfgAAABaT1MvMj4UQ70AAAGYAAAAYGNtYXAViBDZAAABxAAAADZnbHlmz2cGRQAAAeQAAABqaGVhZAHsHgUAAAIQAAAANmhoZWEE0gKgAAACNAAAACRobXR4AAgAAgAAAkgAAAAIbG9jYQAOACgAAAJUAAAACG1heHAAEAAUAAACbAAAABG5hbWUUYGVHAAACgAAAAlRwb3N0PqaBAwAAArQAAABrcHJlcAHrZ3wAAAMMAAAAEQAAACAAAwAAAAQAAUAAAgADAAEAAAAAAAIAAAAAAAEAAQAAAwAAAAAAAQAAAAEAAAAAAAUAAQAAAAAAAgAHADQABAAJAB4AAwABAAAAFAAfACkAAQAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAACAAIAAgAAHgABAAAAAAACAAQAAAABAAEAAwACAAQAAAABAAEAAAAAAAAABAAAAAAABAAAAAAADAAAAAwAAAAMAAAADAAAAAwAAAAAAAQAAAAMAAAAAAAAAAQAAAAMAAAAAAAAAAAAAAAAAAAAAAA==
    `.trim();

    // ✅ الصحيح: على الكائن doc، وليس jsPDF.API
    doc.addFileToVFS('Tajawal-Regular.ttf', tajawalFontBase64);
    doc.addFont('Tajawal-Regular.ttf', 'Tajawal', 'normal');
    doc.setFont('Tajawal');

    // 8) إدراج الصورة
    const imgW = 210;
    const imgH = (canvas.height * imgW) / canvas.width;
    doc.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, imgW, imgH);

    // 9) الحفظ
    const fileName = `سجل_المهارات_${traineeData.name || 'متدرب'}_${traineeData.studentId || currentUser.id}.pdf`;
    doc.save(fileName);

    updateFileOverlay('✅ تم إنشاء ملف PDF!', 'يبدأ التنزيل الآن...');
  } catch (err) {
    console.error('❌ فشل التصدير:', err);
    updateFileOverlay('⚠️ حدث خطأ', err.message || 'فشل إنشاء PDF', true);
  } finally {
    spinner?.classList.add('hidden');
    setTimeout(() => hideFileOverlay(0), 2000);
  }
}



// ✅ نسخة موحّدة: تبني الـ PDF من "المهارات المعتمدة فقط" + تاريخ الورشة
function createPDFHTML(traineeData, skills) {
  // فلترة المعتمد فقط وحساب الإجمالي
  const approved = (skills || []).filter(s => String(s.status).trim() === 'معتمد' && Number(s.hours) > 0);
  const totalHours = approved.reduce((sum, s) => sum + Number(s.hours || 0), 0);

  return `
    <div class="pdf-template" style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 25px; direction: rtl; background: white; min-height: 297mm;">
      <!-- الترويسة -->
      <div style="text-align: center; margin-bottom: 30px; border-bottom: 3px solid #186F65; padding-bottom: 20px;">
        <h1 style="color: #186F65; margin: 0; font-size: 28px; font-weight: bold;">المؤسسة العامة للتدريب التقني والمهني</h1>
        <h2 style="color: #2D3748; margin: 15px 0 0 0; font-size: 22px; font-weight: 600;">سجل المهارات الشخصية</h2>
      </div>
      
      <!-- معلومات المتدرب -->
      <div style="background: #F7FAFC; padding: 20px; border-radius: 12px; margin-bottom: 25px; border-right: 4px solid #186F65;">
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
          <div style="font-size: 16px;"><strong>الاسم:</strong> ${traineeData.name || ''}</div>
          <div style="font-size: 16px;"><strong>رقم الطالب:</strong> ${traineeData.studentId || ''}</div>
          <div style="font-size: 16px;"><strong>التخصص:</strong> ${traineeData.major || traineeData.department || ''}</div>
          <div style="font-size: 16px;"><strong>الفصل الدراسي:</strong> ${traineeData.semester || ''}</div>
          <div style="font-size: 16px;"><strong>السنة التدريبية:</strong> ${traineeData.year || '—'}</div>
        </div>
      </div>
      
      <!-- جدول المهارات (المعتمد فقط) -->
                <div style="margin-bottom:30px;">
            <h3 style="color:#2D3748; margin-bottom:15px; font-size:18px; border-bottom:2px solid #E2E8F0; padding-bottom:8px;">السجل التدريبي</h3>
            <table style="width:100%; border-collapse:collapse; border:1px solid #E2E8F0;">
              <thead>
                <tr style="background:#186F65; color:white;">
                  <th style="padding:12px; border:1px solid #CBD5E0; text-align:center; width:60px;">#</th>
                  <th style="padding:12px; border:1px solid #CBD5E0; text-align:right;">اسم المهارة / الورشة</th>
                  <th style="padding:12px; border:1px solid #CBD5E0; text-align:center; width:100px;">الساعات</th>
                  <th style="padding:12px; border:1px solid #CBD5E0; text-align:center; width:100px;">الحالة</th>
                </tr>
              </thead>
              <tbody>
                ${approved.map((skill, index) => `
                  <tr style="${index % 2 === 0 ? 'background:#F7FAFC;' : 'background:white;'}">
                    <td style="padding:10px; border:1px solid #E2E8F0; text-align:center; font-size:14px;">${index + 1}</td>
                    <td style="padding:10px; border:1px solid #E2E8F0; text-align:right; font-size:14px;"><strong>${skill.name || ''}</strong></td>
                    <td style="padding:10px; border:1px solid #E2E8F0; text-align:center; font-size:14px;">${skill.hours || 0} ساعة</td>
                    <td style="padding:10px; border:1px solid #E2E8F0; text-align:center; font-size:14px;">
                      <span style="color:#186F65; font-weight:bold;">معتمد</span>
                    </td>
                  </tr>
                `).join('')}
                <tr style="background:#EDF2F7; font-weight:bold;">
                  <td style="padding:12px; border:1px solid #CBD5E0; text-align:center;"></td>
                  <td style="padding:12px; border:1px solid #CBD5E0; text-align:right; font-size:15px;">الإجمالي</td>
                  <td style="padding:12px; border:1px solid #CBD5E0; text-align:center; font-size:15px; color:#186F65;">${totalHours} ساعة</td>
                  <td style="padding:12px; border:1px solid #CBD5E0; text-align:center;"></td>
                </tr>
              </tbody>
            </table>
          </div>
      
      <!-- التواقيع -->
      <div style="margin-top: 50px; padding-top: 30px; border-top: 2px solid #186F65;">
        <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 30px; text-align: center;">
          <div>
            <div style="margin-bottom: 60px; font-size: 16px; font-weight: bold; color: #2D3748;">رئيس القسم</div>
            <div style="border-top: 1px solid #CBD5E0; padding-top: 8px; font-size: 14px; color: #4A5568;">
              ${traineeData.headOfDepartment || '_____________'}
            </div>
          </div>
          <div>
            <div style="margin-bottom: 60px; font-size: 16px; font-weight: bold; color: #2D3748;">وكيل شؤون المتدربين</div>
            <div style="border-top: 1px solid #CBD5E0; padding-top: 8px; font-size: 14px; color: #4A5568;">
              ${traineeData.deanOfStudents || '_____________'}
            </div>
          </div>
          <div>
            <div style="margin-bottom: 60px; font-size: 16px; font-weight: bold; color: #2D3748;">العميد</div>
            <div style="border-top: 1px solid #CBD5E0; padding-top: 8px; font-size: 14px; color: #4A5568;">
              ${traineeData.dean || '_____________'}
            </div>
          </div>
        </div>
      </div>

      <div style="margin-top: 40px; text-align: center; color: #718096; font-size: 12px; border-top: 1px solid #E2E8F0; padding-top: 15px;">
        تم إنشاء هذا السجل آلياً عبر نظام سجل المهارات - الكلية التقنية بحقل
      </div>
    </div>
  `;
}


// ✅ إنشاء حاوية PDF
function createPDFContainer() {
  const container = document.createElement('div');
  container.id = 'pdfExportContainer';
  container.style.cssText = `
    position: fixed;
    left: -10000px;
    top: -10000px;
    width: 210mm;
    min-height: 297mm;
    background: white;
    padding: 0;
    margin: 0;
    box-sizing: border-box;
    z-index: -1000;
  `;
  document.body.appendChild(container);
  return container;
}

// ✅ بديل مبسّط باستخدام pdfmake (إذا فشل الحل الأول)
async function exportTraineePDFSimple() {
  try {
    showFileOverlay('📄 جاري إنشاء ملف PDF...', 'قد يستغرق بضع ثوانٍ');

    // جلب البيانات
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=exportTraineeXlsx&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    const data = await res.json();
    
    if (!data.success) throw new Error('تعذر جلب البيانات');
    
    const traineeData = data.traineeData || {};
    const skills = data.skills || [];
    const totalHours = skills.reduce((sum, s) => sum + Number(s.hours || 0), 0);

    // إنشاء محتوى PDF
    const docDefinition = {
      pageSize: 'A4',
      pageMargins: [40, 60, 40, 60],
      content: [
        // الترويسة
        {
          text: 'المؤسسة العامة للتدريب التقني والمهني',
          style: 'header'
        },
        {
          text: 'سجل المهارات الشخصية',
          style: 'subheader'
        },
        
        // معلومات المتدرب
        {
          text: 'معلومات المتدرب',
          style: 'sectionHeader'
        },
        {
          columns: [
            {
              width: '*',
              text: [
                { text: 'الاسم: ', style: 'label' },
                { text: traineeData.name || '', style: 'value' },
                { text: '\nالتخصص: ', style: 'label' },
                { text: traineeData.major || traineeData.department || '', style: 'value' }
              ]
            },
            {
              width: '*',
              text: [
                { text: 'رقم الطالب: ', style: 'label' },
                { text: traineeData.studentId || '', style: 'value' },
                { text: '\nالفصل الدراسي: ', style: 'label' },
                { text: traineeData.semester || '', style: 'value' },
                { text: '\nالسنة التدريبية: ', style: 'label' },
                { text: traineeData.year || '—', style: 'value' }
              ]
            }
          ],
          margin: [0, 0, 0, 20]
        },
        
        // جدول المهارات
        {
          text: 'السجل التدريبي',
          style: 'sectionHeader'
        },
        {
          table: {
            headerRows: 1,
            widths: ['auto', '*', 'auto', 'auto'],
            body: [
              [
                { text: '#', style: 'tableHeader' },
                { text: 'اسم المهارة', style: 'tableHeader' },
                { text: 'الساعات', style: 'tableHeader' },
                { text: 'الحالة', style: 'tableHeader' }
              ],
              ...skills.map((skill, index) => [
                { text: (index + 1).toString(), style: 'tableCell' },
                { text: skill.name || '', style: 'tableCell' },
                { text: `${skill.hours || 0} ساعة`, style: 'tableCell' },
                { text: skill.status || '', style: 'tableCell' }
              ]),
              [
                { text: '', style: 'tableCell' },
                { text: 'الإجمالي', style: 'tableFooter' },
                { text: `${totalHours} ساعة`, style: 'tableFooter' },
                { text: '', style: 'tableCell' }
              ]
            ]
          },
          layout: 'lightHorizontalLines'
        },
        
        // التواقيع
        {
          text: ' ',
          margin: [0, 30, 0, 0]
        },
        {
          columns: [
            {
              width: '*',
              stack: [
                { text: 'رئيس القسم', style: 'signatureLabel' },
                { text: traineeData.headOfDepartment || '_____________', style: 'signatureValue' }
              ],
              alignment: 'center'
            },
            {
              width: '*',
              stack: [
                { text: 'وكيل شؤون المتدربين', style: 'signatureLabel' },
                { text: traineeData.deanOfStudents || '_____________', style: 'signatureValue' }
              ],
              alignment: 'center'
            },
            {
              width: '*',
              stack: [
                { text: 'العميد', style: 'signatureLabel' },
                { text: traineeData.dean || '_____________', style: 'signatureValue' }
              ],
              alignment: 'center'
            }
          ]
        }
      ],
      styles: {
        header: {
          fontSize: 20,
          bold: true,
          color: '#186F65',
          alignment: 'center',
          margin: [0, 0, 0, 10]
        },
        subheader: {
          fontSize: 16,
          bold: true,
          alignment: 'center',
          margin: [0, 0, 0, 30]
        },
        sectionHeader: {
          fontSize: 14,
          bold: true,
          margin: [0, 15, 0, 10],
          color: '#2D3748'
        },
        label: {
          fontSize: 12,
          bold: true,
          color: '#4A5568'
        },
        value: {
          fontSize: 12,
          color: '#2D3748'
        },
        tableHeader: {
          bold: true,
          fontSize: 11,
          color: 'white',
          fillColor: '#186F65',
          alignment: 'center'
        },
        tableCell: {
          fontSize: 10
        },
        tableFooter: {
          bold: true,
          fontSize: 11,
          fillColor: '#F7FAFC'
        },
        signatureLabel: {
          fontSize: 12,
          bold: true,
          margin: [0, 0, 0, 40]
        },
        signatureValue: {
          fontSize: 11
        }
      },
      defaultStyle: {
        font: 'Helvetica',
        alignment: 'right'
      }
    };

    // إنشاء PDF
    pdfMake.createPdf(docDefinition).download(`سجل_المهارات_${traineeData.name || 'متدرب'}_${traineeData.studentId || currentUser.id}.pdf`);
    
    updateFileOverlay('✅ تم إنشاء ملف PDF!', 'يبدأ التنزيل الآن...');
    
  } catch (err) {
    console.error('❌ فشل التصدير:', err);
    updateFileOverlay('⚠️ حدث خطأ', err.message || 'فشل إنشاء PDF', true);
  } finally {
    setTimeout(() => hideFileOverlay(0), 2000);
  }
}


// ✅ تصدير PDF باستخدام pdfmake (دعم عربي أفضل)
async function exportTraineePDFWithPdfMake() {
  try {
    showFileOverlay('📄 جاري إنشاء ملف PDF...', 'قد يستغرق بضع ثوانٍ');

    // جلب البيانات
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=exportTraineeXlsx&userId=${encodeURIComponent(currentUser.id)}`;
    const res = await fetch(url);
    const data = await res.json();
    
    if (!data.success) throw new Error('تعذر جلب البيانات');
    
    const traineeData = data.traineeData || {};
    const skills = data.skills || [];
    const totalHours = skills.reduce((sum, s) => sum + Number(s.hours || 0), 0);

    // تعريف الخطوط العربية
    pdfMake.fonts = {
      Amiri: {
        normal: 'https://cdn.jsdelivr.net/gh/opentypejs/amiri-font/Amiri-Regular.ttf',
        bold: 'https://cdn.jsdelivr.net/gh/opentypejs/amiri-font/Amiri-Bold.ttf',
        italics: 'https://cdn.jsdelivr.net/gh/opentypejs/amiri-font/Amiri-Italic.ttf',
        bolditalics: 'https://cdn.jsdelivr.net/gh/opentypejs/amiri-font/Amiri-BoldItalic.ttf'
      }
    };

    const docDefinition = {
      pageSize: 'A4',
      pageMargins: [40, 60, 40, 60],
      defaultStyle: {
        font: 'Amiri',
        alignment: 'right'
      },
      content: [
        // الترويسة
        {
          text: 'المؤسسة العامة للتدريب التقني والمهني',
          style: 'header'
        },
        {
          text: 'سجل المهارات الشخصية',
          style: 'subheader'
        },
        
        // معلومات المتدرب
        {
          text: 'معلومات المتدرب',
          style: 'sectionHeader'
        },
        {
          columns: [
            {
              width: '*',
              stack: [
                { text: `الاسم: ${traineeData.name || ''}` },
                { text: `التخصص: ${traineeData.major || ''}` }
              ]
            },
            {
              width: '*',
              stack: [
                { text: `رقم الطالب: ${traineeData.studentId || ''}` },
                { text: `الفصل الدراسي: ${traineeData.semester || ''}` },
                { text: `السنة التدريبية: ${traineeData.year || '—'}` }
              ]
            }
          ],
          margin: [0, 0, 0, 15]
        },
        
        // جدول المهارات
        {
          text: 'السجل التدريبي',
          style: 'sectionHeader'
        },
        {
          table: {
            headerRows: 1,
            widths: ['auto', '*', 'auto', 'auto'],
            body: [
              [
                { text: '#', style: 'tableHeader' },
                { text: 'اسم المهارة', style: 'tableHeader' },
                { text: 'الساعات', style: 'tableHeader' },
                { text: 'الحالة', style: 'tableHeader' }
              ],
              ...skills.map((skill, index) => [
                { text: (index + 1).toString(), style: 'tableCell' },
                { text: skill.name || '', style: 'tableCell' },
                { text: `${skill.hours || 0} ساعة`, style: 'tableCell' },
                { text: skill.status || '', style: 'tableCell' }
              ]),
              [
                { text: '', style: 'tableCell' },
                { text: 'الإجمالي', style: 'tableFooter' },
                { text: `${totalHours} ساعة`, style: 'tableFooter' },
                { text: '', style: 'tableCell' }
              ]
            ]
          },
          layout: {
            hLineWidth: () => 1, vLineWidth: () => 1,
            hLineColor: () => '#cccccc', vLineColor: () => '#cccccc',
            paddingLeft: () => 8, paddingRight: () => 8,
            paddingTop: () => 4, paddingBottom: () => 4
          }
        },
        
        // التواقيع
        {
          text: ' ',
          margin: [0, 30, 0, 0]
        },
        {
          columns: [
            {
              width: '*',
              stack: [
                { text: 'رئيس القسم', style: 'signatureLabel' },
                { text: traineeData.headOfDepartment || '_____________', style: 'signatureValue' }
              ],
              alignment: 'center'
            },
            {
              width: '*',
              stack: [
                { text: 'وكيل شؤون المتدربين', style: 'signatureLabel' },
                { text: traineeData.deanOfStudents || '_____________', style: 'signatureValue' }
              ],
              alignment: 'center'
            },
            {
              width: '*',
              stack: [
                { text: 'العميد', style: 'signatureLabel' },
                { text: traineeData.dean || '_____________', style: 'signatureValue' }
              ],
              alignment: 'center'
            }
          ]
        }
      ],
      styles: {
        header: {
          fontSize: 20,
          bold: true,
          color: '#186F65',
          alignment: 'center',
          margin: [0, 0, 0, 10]
        },
        subheader: {
          fontSize: 16,
          bold: true,
          alignment: 'center',
          margin: [0, 0, 0, 30]
        },
        sectionHeader: {
          fontSize: 14,
          bold: true,
          margin: [0, 15, 0, 10],
          color: '#333'
        },
        tableHeader: {
          bold: true,
          fontSize: 12,
          color: 'white',
          fillColor: '#186F65',
          alignment: 'center'
        },
        tableCell: {
          fontSize: 10,
          margin: [0, 4, 0, 4]
        },
        tableFooter: {
          bold: true,
          fontSize: 11,
          fillColor: '#f0f0f0'
        },
        signatureLabel: {
          fontSize: 12,
          bold: true,
          margin: [0, 0, 0, 40]
        },
        signatureValue: {
          fontSize: 11,
          margin: [0, 10, 0, 0]
        }
      }
    };

    // إنشاء PDF
    pdfMake.createPdf(docDefinition).download(`سجل_المهارات_${traineeData.name}_${traineeData.studentId}.pdf`);
    
    updateFileOverlay('✅ تم إنشاء ملف PDF!', 'يبدأ التنزيل الآن...');
    
  } catch (err) {
    console.error('❌ فشل التصدير:', err);
    updateFileOverlay('⚠️ حدث خطأ', err.message || 'فشل إنشاء PDF', true);
  } finally {
    setTimeout(() => hideFileOverlay(0), 2000);
  }
}

// ✅ صحيحة: نسخة موحّدة لا تتكرر، تختار أفضل مسار تلقائيًا
// 1) تفضّل html2canvas + jsPDF (أفضل للعربية)
// 2) إن تعذّر، تستخدم pdfmake
// 3) إن لم تتوفر مكتبات PDF، تعرض تنبيه واضح
async function exportSkillsPDF() {
  if (typeof html2canvas !== 'undefined' && window.jspdf) {
    await exportTraineePDF();            // يستخدم createPDFHTML + الخطوط + jsPDF
  } else if (typeof pdfMake !== 'undefined') {
    await exportTraineePDFSimple();      // بديل pdfmake مع جداول عربية
  } else {
    alert('⚠️ لم يتم تحميل مكتبات PDF. تأكد من إضافة html2canvas و jsPDF أو pdfmake في index.html');
    console.error('مكتبات PDF غير محمّلة:', {
      html2canvas: typeof html2canvas,
      jspdf: typeof window.jspdf,
      pdfmake: typeof pdfMake
    });
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
  upgradeIcons(document.getElementById('workshopDetailsModal'));

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

// ===============================================
// ⚡ نظام الاعتماد الجماعي للمدرب
// ===============================================

// ملء قائمة الورش في select الاعتماد الجماعي
function populateBulkWorkshopSelect() {
  const select = document.getElementById('bulkWorkshopSelect');
  if (!select || !window.trainerWorkshops) return;
  
  // فلترة الورش النشطة/المتاحة فقط
  const activeWorkshops = (window.trainerWorkshops || []).filter(w => 
    w.status === 'نشط' || w.status === 'متاح'
  );
  
  select.innerHTML = '<option value="">-- اختر ورشة --</option>' +
    activeWorkshops.map(w => 
      `<option value="${w.id}">${w.name} (${w.date})</option>`
    ).join('');
}

// اعتماد جماعي لجميع السجلات المعلقة
async function bulkApproveAll(action) {
  const actionText = action === 'approve' ? 'اعتماد حضور' : 
                     action === 'noshow' ? 'تعليم كـ "لم يحضر" لـ' : 
                     'إعادة لحالة معلق لـ';
  
  const confirm = window.confirm(
    `هل أنت متأكد من ${actionText} جميع السجلات المعلقة؟\n\n` +
    `⚠️ هذا الإجراء سيؤثر على جميع المتدربين في كل ورشك.`
  );
  
  if (!confirm) return;
  
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  
  try {
    showFileOverlay('⚡ جاري المعالجة الجماعية...', 'قد يستغرق بضع ثوانٍ');
    
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'bulkApproveAttendance',
        trainerId: currentUser.id,
        workshopId: null, // null = كل الورش
        action: action
      })
    });
    
    const data = await res.json();
    
    if (data?.success) {
      updateFileOverlay(
        '✅ تمت المعالجة بنجاح', 
        `تم ${actionText} ${data.processedCount || 0} سجل`
      );
      
      setTimeout(async () => {
        hideFileOverlay(0);
        await loadTrainerData(); // إعادة تحميل البيانات
        alert(`✅ تم ${actionText} ${data.processedCount || 0} سجل بنجاح!`);
      }, 2000);
    } else {
      throw new Error(data?.message || 'فشلت العملية');
    }
  } catch (err) {
    console.error('خطأ في الاعتماد الجماعي:', err);
    updateFileOverlay('❌ فشلت العملية', err.message, true);
    setTimeout(() => hideFileOverlay(0), 3000);
  } finally {
    spinner?.classList.add('hidden');
  }
}

// اعتماد جماعي لورشة محددة
async function bulkApproveByWorkshop(action) {
  const select = document.getElementById('bulkWorkshopSelect');
  const workshopId = select?.value;
  
  if (!workshopId) {
    alert('⚠️ يرجى اختيار ورشة أولاً');
    return;
  }
  
  const workshopName = select.options[select.selectedIndex].text;
  const actionText = action === 'approve' ? 'اعتماد حضور' : 'تعليم كـ "لم يحضر" لـ';
  
  const confirm = window.confirm(
    `هل أنت متأكد من ${actionText} جميع المتدربين في:\n\n` +
    `📋 ${workshopName}\n\n` +
    `⚠️ سيتم تطبيق هذا على جميع السجلات المعلقة في هذه الورشة.`
  );
  
  if (!confirm) return;
  
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  
  try {
    showFileOverlay('⚡ جاري المعالجة...', `معالجة الورشة: ${workshopName}`);
    
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'bulkApproveAttendance',
        trainerId: currentUser.id,
        workshopId: workshopId,
        action: action
      })
    });
    
    const data = await res.json();
    
    if (data?.success) {
      updateFileOverlay(
        '✅ تمت المعالجة بنجاح', 
        `تم ${actionText} ${data.processedCount || 0} متدرب`
      );
      
      setTimeout(async () => {
        hideFileOverlay(0);
        await loadTrainerData(); // إعادة تحميل البيانات
        alert(`✅ تم ${actionText} ${data.processedCount || 0} متدرب في الورشة!`);
        
        // إعادة تعيين القائمة
        select.value = '';
      }, 2000);
    } else {
      throw new Error(data?.message || 'فشلت العملية');
    }
  } catch (err) {
    console.error('خطأ في الاعتماد الجماعي:', err);
    updateFileOverlay('❌ فشلت العملية', err.message, true);
    setTimeout(() => hideFileOverlay(0), 3000);
  } finally {
    spinner?.classList.add('hidden');
  }
}

// تعديل renderTrainerWorkshops لملء select الورش
const originalRenderTrainerWorkshops = renderTrainerWorkshops;
renderTrainerWorkshops = function(workshops) {
  originalRenderTrainerWorkshops(workshops);
  populateBulkWorkshopSelect(); // ملء قائمة الورش للاعتماد الجماعي
};

// --- ضع هذه الدالة مرّة واحدة أعلى الملف (مساعدة آمنة للتنزيل) ---
function triggerBlobDownload(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;

  // Safari iOS/macOS قد يتجاهل download مع blob إذا لم يكن العنصر في الـ DOM
  document.body.appendChild(a);

  // إذا المتصفح لا يدعم download (بعض Safari قد يعيده undefined)
  if (typeof a.download === 'undefined') {
    window.open(url, '_blank'); // فتح الصورة في تبويب — يقدر المستخدم يحفظها
  } else {
    a.click();
  }

  // تنظيف
  setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 0);
}
 

// ====== 🎓 نظام شهادات PDF احترافي مع دعم العربية الكامل ======

// ✅ المسار الصحيح للخلفية
const CERT_BG_URL = 'templates/cert_bg_a4.png'; // ← غيّرها لمسار قالبك إن لزم

// ✅ أبعاد الشهادة (A4 Landscape)
const CERT_WIDTH = 297; // mm
const CERT_HEIGHT = 210; // mm

// تحكم في المسافات الرأسية حول الاسم
// مسافات قابلة للتعديل بدقّة (بالملِّيمتر)
const PREAMBLE_TOP_SHIFT_MM   = 15;  // نزول فقرة "تشهد المؤسسة..." عن أعلى الصفحة
const GAP_PREAMBLE_TO_NAME_MM = 4;   // المسافة بين "بأن المتدرب" واسم المتدرب
const GAP_NAME_TO_SKILL_MM    = 2.5; // المسافة بين اسم المتدرب و"قد أتم بنجاح"
const CERT_SPACING_ABOVE_NAME_MM = 0.5; // المسافة بين "بأن المتدرب/ة" واسم المتدرب
const CERT_NAME_BOTTOM_MM = 0;        // المسافة تحت اسم المتدرب قبل نص المهارة
const SKILL_TOP_MARGIN_MM = 2;          // جديد: مسافة صغيرة جداً قبل "قد أتم بنجاح"
const CERT_NAME_UNDERLINE_OFFSET_MM = 1.5; // بُعد خط التسطير (اختياري لتقليل الفراغ البصري)


// ارفع النص للأعلى بمقدار 3 سم (30mm)
const CERT_TEXT_TOP_MM = 10; // كان 60mm، الآن 30mm = أعلى بمقدار 3cm


const SIGN_LINE_COLOR = '#000';     // لون الخط (أسود)
const SIGN_LINE_THICK = 2;          // سماكة الخط px
const SIGN_GAP_ABOVE_LINE_MM = 3;   // مسافة فوق الخط (بين الاسم والخط)
const SIGN_ROLE_SIZE_PT = 16;       // حجم نص المسمّى
const SIGN_NAME_SIZE_PT = 11;       // حجم اسم الموقّع
const SHOW_SIGN_LINE = false; // ← خلّيه false لحذف الخط



// ✅ إنشاء HTML للشهادة مع التصميم الكامل
function createCertificateHTML({
  traineeName,
  skillName,
  hours,
  dateText,
  serial,
  signatures,
  bgImage
}) {
  // تحديد اتجاه الصفحة
  const pageDir =
    (document?.documentElement?.dir || document?.dir || 'rtl').toLowerCase();
  const isRTL = pageDir === 'rtl' ||
                /[\u0600-\u06FF]/.test(String(traineeName) + String(skillName));

  // زاوية الرقم حسب الاتجاه
  const serialCornerStyle = isRTL
    ? 'right:8mm; left:auto; text-align:right; direction:rtl;'
    : 'left:8mm; right:auto; text-align:left; direction:ltr;';

  // إزاحة إضافية 2 سم للنص "تشهد المؤسسة..."
  const PREAMBLE_EXTRA_OFFSET_MM = 15;

  return `
    <div id="certificateContainer" style="
      width:${CERT_WIDTH}mm;
      height:${CERT_HEIGHT}mm;
      position:relative;
      background:white;
      font-family:'Cairo','Tajawal','Segoe UI',Tahoma,sans-serif;
      direction:${isRTL ? 'rtl' : 'ltr'};
      overflow:hidden;
    ">
      <!-- الخلفية -->
      <img src="${bgImage}" style="position:absolute; inset:0; width:100%; height:100%; object-fit:cover;" />

      <!-- المحتوى -->
      <div style="
        position:absolute; inset:0;
        display:flex; flex-direction:column; justify-content:center; align-items:center;
        text-align:center; padding:40mm 30mm; box-sizing:border-box;
      ">

        <!-- نص الشهادة (مرفوع/مُنزل حسب الطلب) -->
        <div style="margin-top:${CERT_TEXT_TOP_MM}mm;">
        <p style="
          margin:${PREAMBLE_TOP_SHIFT_MM}mm 0 ${GAP_PREAMBLE_TO_NAME_MM}mm 0;
          font-size:16pt; font-weight:600; line-height:1.6;
          color:#ffffff; text-shadow:0 1px 2px rgba(0,0,0,.6);
        ">
          تشهد المؤسسة العامة للتدريب التقني والمهني<br>
          الكلية التقنية بحقل<br>
          بأن المتدرب
        </p>

        <!-- اسم المتدرب -->
        <h1 style="
          font-size:28pt; font-weight:800; line-height:1.1;
          margin:0 0 ${GAP_NAME_TO_SKILL_MM}mm 0;
          color:#ffffff; text-shadow:0 2px 3px rgba(0,0,0,.65);
          text-decoration:underline; text-decoration-color:#ffffff;
          text-underline-offset:${CERT_NAME_UNDERLINE_OFFSET_MM}mm;
        ">
          ${traineeName}
        </h1>

        <!-- "قد أتم بنجاح..." + اسم المهارة -->
        <p style="
          font-size:18pt; margin:0; font-weight:600;
          color:#E6F6F3; text-shadow:0 1px 2px rgba(0,0,0,.5);
        ">
          قد أتم بنجاح الدورة التدريبية<br>
          <span style="
            font-size:22pt; font-weight:800; display:inline-block;
            margin-top:${SKILL_TOP_MARGIN_MM}mm; color:#ffffff;
          ">
            ${skillName}
          </span>
        </p>


          <!-- الساعات والتاريخ -->
          <div style="
            display:flex; justify-content:center; gap:20mm; margin:8mm 0; font-size:14pt;
            color:#F1FCFA; text-shadow:0 1px 2px rgba(0,0,0,.45);
          ">
            <div><strong>عدد الساعات:</strong> ${hours} ساعة تدريبية</div>
            <div><strong>التاريخ:</strong> ${dateText}</div>
          </div>
        </div>

        <!-- التواقيع (رئيس القسم → وكيل شؤون المتدربين → العميد) -->
        <div style="
          position:absolute; bottom:20mm; left:0; right:0;
          display:flex; justify-content:space-around; padding:0 30mm;
        ">
          <!-- رئيس القسم -->
          <div style="text-align:center; min-width:60mm;">
            <div style="font-size:${SIGN_ROLE_SIZE_PT}pt; font-weight:700; color:#ffffff; margin-bottom:8mm; text-shadow:0 1px 2px rgba(0,0,0,.5);">
              رئيس القسم
            </div>
            <div style="font-size:${SIGN_NAME_SIZE_PT}pt; font-weight:600; color:#F1FCFA; text-shadow:0 1px 2px rgba(0,0,0,.4);">
              ${signatures.hod || '_____________'}
            </div>
          </div>

          <!-- وكيل شؤون المتدربين -->
          <div style="text-align:center; min-width:60mm;">
            <div style="font-size:${SIGN_ROLE_SIZE_PT}pt; font-weight:700; color:#ffffff; margin-bottom:8mm; text-shadow:0 1px 2px rgba(0,0,0,.5);">
              وكيل شؤون المتدربين
            </div>
            <div style="font-size:${SIGN_NAME_SIZE_PT}pt; font-weight:600; color:#F1FCFA; text-shadow:0 1px 2px rgba(0,0,0,.4);">
              ${signatures.deanStd || '_____________'}
            </div>
          </div>

          <!-- العميد -->
          <div style="text-align:center; min-width:60mm;">
            <div style="font-size:${SIGN_ROLE_SIZE_PT}pt; font-weight:700; color:#ffffff; margin-bottom:8mm; text-shadow:0 1px 2px rgba(0,0,0,.5);">
              العميد
            </div>
            <div style="font-size:${SIGN_NAME_SIZE_PT}pt; font-weight:600; color:#F1FCFA; text-shadow:0 1px 2px rgba(0,0,0,.4);">
              ${signatures.dean || '_____________'}
            </div>
          </div>
        </div>

        <!-- الرقم التسلسلي (يسار) -->
        <div style="
          position:absolute; bottom:6mm; left:8mm; right:auto;
          text-align:left; direction:ltr;
          font-size:10pt; font-weight:700; letter-spacing:.4px;
          color:#ffffff; text-shadow:0 1px 2px rgba(0,0,0,.6); opacity:.95;
          font-family:monospace;
        ">
          ${serial}
        </div>

      </div>
    </div>
  `;
}


// ✅ تحويل صورة إلى Data URL
async function imageToDataURL(url) {
  return new Promise((resolve, reject) => {
    console.log('🔄 بدء تحميل الصورة:', url);
    
    const xhr = new XMLHttpRequest();
    
    xhr.timeout = 15000; // 15 ثانية timeout
    
    xhr.onload = function() {
      if (xhr.status === 200) {
        const reader = new FileReader();
        reader.onloadend = () => {
          console.log('✅ تم تحميل الصورة بنجاح');
          resolve(reader.result);
        };
        reader.onerror = (error) => {
          console.error('❌ خطأ في FileReader:', error);
          reject(new Error('فشل قراءة ملف الصورة'));
        };
        reader.readAsDataURL(xhr.response);
      } else {
        reject(new Error(`فشل تحميل الصورة: ${xhr.status} ${xhr.statusText}`));
      }
    };
    
    xhr.onerror = () => {
      console.error('❌ خطأ شبكة في تحميل الصورة');
      reject(new Error('خطأ في الاتصال بالشبكة'));
    };
    
    xhr.ontimeout = () => {
      console.error('⏱️ انتهت مهلة تحميل الصورة');
      reject(new Error('انتهت مهلة تحميل الصورة. يرجى المحاولة مرة أخرى'));
    };
    
    xhr.open('GET', url);
    xhr.responseType = 'blob';
    xhr.send();
  });
}

// ✅ الدالة الرئيسية: توليد شهادة PDF
async function generateCertificatePDF({
  traineeName,
  skillName,
  hours,
  dateText,
  serial,
  signatures
}) {
  const spinner = document.getElementById('loadingSpinner');
  let container = null;
  
  try {
    // 1) التحقق من المكتبات
    console.log('🔍 فحص المكتبات المطلوبة...');
    
    if (typeof html2canvas === 'undefined') {
      console.error('❌ html2canvas غير محملة');
      throw new Error('مكتبة html2canvas غير محملة. يرجى إعادة تحميل الصفحة.');
    }
    if (typeof window.jspdf === 'undefined') {
      console.error('❌ jsPDF غير محملة');
      throw new Error('مكتبة jsPDF غير محملة. يرجى إعادة تحميل الصفحة.');
    }
    
    console.log('✅ المكتبات متوفرة');

    spinner?.classList.remove('hidden');
    updateFileOverlay('📄 جاري إنشاء الشهادة...', 'تحميل القالب');

    // 2) تحميل صورة الخلفية مع timeout
    console.log('📥 تحميل صورة الخلفية من:', CERT_BG_URL);
    updateFileOverlay('🎨 تجهيز التصميم...', 'تحميل الخلفية');
    
    let bgDataURL;
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 15000); // 15 ثانية
      
      bgDataURL = await imageToDataURL(CERT_BG_URL);
      clearTimeout(timeoutId);
      
      console.log('✅ تم تحميل الخلفية بنجاح');
    } catch (bgError) {
      console.error('❌ فشل تحميل الخلفية:', bgError);
      // استخدام خلفية بيضاء كبديل
      bgDataURL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';
      console.warn('⚠️ سيتم استخدام خلفية بيضاء بديلة');
    }

    // 3) إنشاء HTML المؤقت
    updateFileOverlay('✍️ كتابة البيانات...', 'إضافة المعلومات');
    console.log('📝 إنشاء HTML للشهادة...');
    
    const certHTML = createCertificateHTML({
      traineeName,
      skillName,
      hours,
      dateText,
      serial,
      signatures,
      bgImage: bgDataURL
    });

    // 4) حقن HTML في container مخفي
    console.log('🔧 إنشاء الحاوية المؤقتة...');
    
    container = document.getElementById('pdfCertContainer');
    if (!container) {
      container = document.createElement('div');
      container.id = 'pdfCertContainer';
      container.style.cssText = `
        position: fixed;
        left: -10000px;
        top: -10000px;
        z-index: -9999;
      `;
      document.body.appendChild(container);
    }
    container.innerHTML = certHTML;

    // 5) انتظار تحميل الخطوط والصور
    updateFileOverlay('⏳ تحميل الخطوط...', 'يرجى الانتظار قليلاً');
    console.log('⏳ انتظار تحميل الخطوط...');
    
    try {
      if (document.fonts && document.fonts.ready) {
        await Promise.race([
          document.fonts.ready,
          new Promise(resolve => setTimeout(resolve, 3000)) // timeout 3 ثواني
        ]);
      } else {
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
      console.log('✅ الخطوط جاهزة');
    } catch (fontError) {
      console.warn('⚠️ تعذر انتظار الخطوط:', fontError);
    }
    
    // مهلة إضافية للتأكد من رسم كل شيء
    await new Promise(resolve => setTimeout(resolve, 800));

    // 6) تحويل HTML إلى صورة
    updateFileOverlay('📸 التقاط الشهادة...', 'تحويل إلى صورة عالية الدقة');
    console.log('📸 بدء html2canvas...');
    
    const element = document.getElementById('certificateContainer');
    
    if (!element) {
      throw new Error('لم يتم العثور على عنصر الشهادة في الصفحة');
    }
    
    const canvas = await html2canvas(element, {
      scale: 3, // دقة عالية
      useCORS: true,
      allowTaint: false,
      backgroundColor: '#ffffff',
      logging: false,
      width: element.offsetWidth,
      height: element.offsetHeight,
      onclone: (clonedDoc) => {
        console.log('✅ تم نسخ المستند للمعالجة');
      }
    });
    
    console.log('✅ تم إنشاء Canvas بنجاح:', canvas.width, 'x', canvas.height);

    // 7) إنشاء PDF
    updateFileOverlay('📦 إنشاء ملف PDF...', 'التحضير للتنزيل');
    console.log('📦 إنشاء ملف PDF...');
    
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4',
      compress: true
    });

    const imgData = canvas.toDataURL('image/jpeg', 0.95); // استخدام JPEG للضغط
    pdf.addImage(imgData, 'JPEG', 0, 0, CERT_WIDTH, CERT_HEIGHT);
    
    console.log('✅ تم إضافة الصورة للـ PDF');

    // 8) التنزيل
    updateFileOverlay('💾 حفظ الملف...', 'جاري التنزيل');
    
    const fileName = `شهادة_${traineeName.replace(/[\\/:*?"<>|]/g, '_')}_${skillName.replace(/[\\/:*?"<>|]/g, '_')}.pdf`;
    console.log('💾 حفظ الملف:', fileName);
    
    pdf.save(fileName);

    updateFileOverlay('✅ تم إنشاء الشهادة!', 'اكتمل التنزيل بنجاح');
    
    console.log('🎉 تم إنشاء الشهادة بنجاح!');
    
    setTimeout(() => {
      hideFileOverlay(0);
      if (container && container.parentNode) {
        container.remove();
      }
    }, 2000);

  } catch (error) {
    console.error('❌ فشل إنشاء الشهادة:', error);
    console.error('تفاصيل الخطأ:', {
      message: error.message,
      stack: error.stack
    });
    
    updateFileOverlay('⚠️ فشل إنشاء الشهادة', error.message, true);
    setTimeout(() => hideFileOverlay(0), 4000);
    throw error;
  } finally {
    spinner?.classList.add('hidden');
    
    // تنظيف الحاوية المؤقتة
    if (container && container.parentNode) {
      setTimeout(() => {
        try {
          container.remove();
        } catch (e) {
          console.warn('تعذر إزالة الحاوية:', e);
        }
      }, 3000);
    }
  }
}

// ✅ الدالة المستدعاة من الواجهة
async function downloadCertificatePDF(skillId) {
  const spinner = document.getElementById('loadingSpinner');
  
  try {
    // 1) إظهار الانتظار فوراً
    spinner?.classList.remove('hidden');
    showFileOverlay('🎓 جاري تحضير الشهادة...', 'يرجى الانتظار');
    
    console.log('📋 بدء عملية إنشاء الشهادة للمهارة:', skillId);
    
    // 2) البحث عن المهارة
    updateFileOverlay('🔍 جاري البحث عن المهارة...', 'قراءة البيانات');
    
    const skill = (window.traineeSkillsCache || []).find(s => 
      String(s.id) === String(skillId)
    );
    
    if (!skill) {
      console.error('❌ لم يتم العثور على المهارة:', skillId);
      throw new Error('تعذر العثور على بيانات المهارة');
    }
    
    console.log('✅ تم العثور على المهارة:', skill.name);

    // 3) التحقق من الصلاحية
    const hours = Number(skill.hours || 0);
    const status = String(skill.status || '').trim();
    
    if (status !== 'معتمد') {
      updateFileOverlay('⚠️ تنبيه', 'الشهادة متاحة فقط للمهارات المعتمدة', true);
      setTimeout(() => {
        hideFileOverlay(0);
        alert('⚠️ الشهادة متاحة فقط للمهارات المعتمدة');
      }, 2000);
      return;
    }
    
    if (hours <= 0) {
      updateFileOverlay('⚠️ تنبيه', 'عدد الساعات يجب أن يكون أكبر من صفر', true);
      setTimeout(() => {
        hideFileOverlay(0);
        alert('⚠️ عدد الساعات يجب أن يكون أكبر من صفر');
      }, 2000);
      return;
    }

    // 4) جمع البيانات الأساسية
    updateFileOverlay('📝 جمع البيانات...', 'قراءة معلومات المتدرب');
    
    const traineeName = currentUser?.name || 'المتدرب';
    const skillName = skill.name || 'مهارة';
    const dateText = skill.date || new Date().toLocaleDateString('ar-SA');
    
    console.log('📊 بيانات الشهادة:', { traineeName, skillName, hours, dateText });

    // 5) جلب التواقيع مع timeout
    updateFileOverlay('🔐 جلب التواقيع...', 'الاتصال بالخادم');
    
    let hod = '';
    let deanStd = '';
    let dean = '';
    
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 ثواني timeout
      
      const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=exportTraineeXlsx&userId=${encodeURIComponent(currentUser.id)}`;
      const res = await fetch(url, { signal: controller.signal });
      clearTimeout(timeoutId);
      
      const data = await res.json();
      
      if (data?.success && data.traineeData) {
        updateFileOverlay('✅ تم استلام البيانات', 'معالجة التواقيع');
        
        const t = data.traineeData;
        t.major = t.major || t.department || currentUser?.department || '';
        await resolveSignaturesOnlineStrict(t);
        
        hod = t.headOfDepartment || '';
        deanStd = t.deanOfStudents || '';
        dean = t.dean || '';
        
        console.log('✅ تم جلب التواقيع:', { hod, deanStd, dean });
      }
    } catch (e) {
      console.warn('⚠️ تعذر جلب التواقيع من الخادم:', e.message);
      // المتابعة بدون توقيعات
    }

    // 6) رقم تسلسلي
    const now = new Date();
    const dateStamp = now.toISOString().slice(0, 10).replace(/-/g, '');
    const studentKey = (currentUser?.studentId || currentUser?.id || 'USR')
      .toString()
      .replace(/\W+/g, '')
      .slice(-6)
      .toUpperCase();
    const skillKey = String(skillId).slice(-4).toUpperCase();
    const certSerial = `CERT-${studentKey}-${skillKey}-${dateStamp}`;

    // 7) إنشاء الشهادة
    updateFileOverlay('🎨 إنشاء الشهادة...', 'رسم التصميم');
    
    await generateCertificatePDF({
      traineeName,
      skillName,
      hours,
      dateText,
      serial: certSerial,
      signatures: { hod, deanStd, dean }
    });
    
    console.log('✅ تم إنشاء الشهادة بنجاح');
    
  } catch (error) {
    console.error('❌ خطأ في downloadCertificatePDF:', error);
    updateFileOverlay('⚠️ حدث خطأ', error.message || 'فشل إنشاء الشهادة', true);
    setTimeout(() => {
      hideFileOverlay(0);
      alert('❌ ' + error.message);
    }, 2000);
  } finally {
    spinner?.classList.add('hidden');
  }
}

// ✅ ضعها خارج downloadCertificate (كانت متداخلة بالخطأ)
function viewDepartmentDetails(departmentName){
  alert('سيتم فتح تفاصيل قسم: ' + departmentName);
}



// ---- عند تحميل الصفحة: استعادة الجلسة مع التطبيع ----
window.onload = function () {
  const saved = localStorage.getItem('currentUser');
  if (saved) {
    try {
      currentUser = JSON.parse(saved);
      currentUser.userType = normalizeUserType(currentUser.userType);
      localStorage.setItem('currentUser', JSON.stringify(currentUser));
      redirectToDashboard(currentUser.userType);
      try{ upgradeIcons(); }catch{}
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
  if (!currentUser || currentUser.userType !== CONFIG.USER_TYPES.TRAINER) {
    alert('هذه الشاشة مخصصة للمدرب فقط.');
    return;
  }

  const modal = document.getElementById('workshopDetailsModal');
  const spinner = document.getElementById('loadingSpinner');
  
  const ws = (window.trainerWorkshops || []).find(w => String(w.id) === String(workshopId));
  if (!ws) {
    alert('تعذر العثور على الورشة.');
    return;
  }

  // ✅ فتح المودال مع بيانات الورشة
  document.getElementById('wDetName').textContent = ws.name;
  document.getElementById('wDetDate').textContent = ws.date || '-';
  document.getElementById('wDetLocation').textContent = ws.location || '-';
  document.getElementById('wDetCapacity').textContent = `${ws.registered || 0}/${ws.capacity || 0}`;
  
  const statusEl = document.getElementById('wDetStatus');
  statusEl.textContent = ws.status || '-';
  statusEl.className = 'workshop-badge ' + 
    ((ws.status === 'نشط' || ws.status === 'متاح') ? 'badge-available' : 'badge-completed');
  
  modal.classList.add('active');

  // ✅ جلب قائمة المتدربين
  spinner?.classList.remove('hidden');
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=getWorkshopDetails&workshopId=${encodeURIComponent(workshopId)}`;
    const res = await fetch(url);
    const data = await res.json();

    let participants = [];
    if (data?.success && Array.isArray(data.participants)) {
      participants = data.participants.map(p => ({
        id: p.attendanceId,
        traineeId: p.traineeId,
        traineeName: p.traineeName,
        status: p.status
      }));
    }
    
    renderWorkshopParticipants(document.getElementById('wDetParticipants'), participants);
  } catch (e) {
    console.error('خطأ في جلب تفاصيل الورشة:', e);
    document.getElementById('wDetParticipants').innerHTML = 
      '<p style="color:var(--error);">تعذر جلب قائمة المتدربين</p>';
  } finally {
    spinner?.classList.add('hidden');
  }
}


// ===============================================
// 📤 نظام الشهادات الخارجية
// ===============================================

function openUploadCertModal() {
  document.getElementById('uploadCertModal')?.classList.add('active');
}

function closeUploadCertModal() {
  document.getElementById('uploadCertModal')?.classList.remove('active');
  document.getElementById('certCourseName').value = '';
  document.getElementById('certHours').value = '';
  document.getElementById('certFile').value = '';
}

// رفع ملف إلى Google Drive (عبر base64)
async function uploadFileToGoogleDrive(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(',')[1];
      resolve({
        name: file.name,
        mimeType: file.type,
        content: base64
      });
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function submitExternalCert(event) {
  event.preventDefault();
  
  const courseName = document.getElementById('certCourseName').value.trim();
  const hours = Number(document.getElementById('certHours').value);
  const fileInput = document.getElementById('certFile');
  const file = fileInput.files[0];
  
  if (!file) {
    return alert('يرجى اختيار ملف الشهادة');
  }
  
  // التحقق من حجم الملف (5MB)
  if (file.size > 5242880) {
    return alert('حجم الملف كبير جداً. الحد الأقصى 5MB');
  }
  
  const spinner = document.getElementById('loadingSpinner');
  spinner?.classList.remove('hidden');
  
  try {
    showFileOverlay('📤 جاري رفع الشهادة...', 'قد يستغرق بضع ثوانٍ');
    
    // تحويل الملف إلى base64
    const fileData = await uploadFileToGoogleDrive(file);
    
    // إرسال البيانات إلى GAS
    const res = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'submitExternalCertificate',
        userId: currentUser.id,
        courseName,
        hours,
        fileName: fileData.name,
        mimeType: fileData.mimeType,
        fileContent: fileData.content
      })
    });
    
    const data = await res.json();
    
    if (data?.success) {
      updateFileOverlay('✅ تم الرفع بنجاح', 'سيتم اعتمادها من رئيس القسم');
      setTimeout(() => hideFileOverlay(0), 2000);
      
      closeUploadCertModal();
      await loadTraineeData(); // إعادة تحميل البيانات
      alert('✅ تم رفع الشهادة بنجاح!\nسيتم مراجعتها واعتمادها من قبل رئيس القسم.');
    } else {
      throw new Error(data?.message || 'فشل رفع الشهادة');
    }
  } catch (err) {
    console.error('خطأ في رفع الشهادة:', err);
    updateFileOverlay('❌ فشل الرفع', err.message, true);
    setTimeout(() => hideFileOverlay(0), 3000);
  } finally {
    spinner?.classList.add('hidden');
  }
}

// جلب وعرض الشهادات الخارجية للمتدرب
async function loadExternalCerts() {
  if (!currentUser || currentUser.userType !== CONFIG.USER_TYPES.TRAINEE) return;
  
  try {
    const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=listExternalCertificates&department=${encodeURIComponent(currentUser.department || '')}`;
    const res = await fetch(url);
    const data = await res.json();
    
    if (data?.success) {
      const myCerts = (data.certificates || []).filter(c => 
        String(c.userId) === String(currentUser.id)
      );
      renderExternalCerts(myCerts);
    }
  } catch (e) {
    console.error('خطأ في جلب الشهادات الخارجية:', e);
  }
}

function renderExternalCerts(certs) {
  const container = document.getElementById('externalCertsList');
  if (!container) return;
  
  if (!certs || !certs.length) {
    container.innerHTML = '<p style="text-align:center;color:var(--tvtc-text-muted);">لم يتم رفع أي شهادات خارجية بعد.</p>';
    return;
  }
  
  const rows = certs.map(c => {
    const badgeClass = c.status === 'معتمد' ? 'badge-completed' : 
                       c.status === 'مرفوض' ? 'badge-error' : 'badge-pending';
    
    return `
      <tr>
        <td><strong>${c.courseName}</strong></td>
        <td>${c.hours} ساعة</td>
        <td><span class="workshop-badge ${badgeClass}">${c.status}</span></td>
        <td>
          ${c.fileUrl ? `<a href="${c.fileUrl}" target="_blank" class="btn btn-outline btn-small">📎 عرض</a>` : '-'}
        </td>
      </tr>
    `;
  }).join('');
  
  container.innerHTML = `
    <table>
      <thead>
        <tr><th>اسم الدورة</th><th>الساعات</th><th>الحالة</th><th>الملف</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

// تعديل loadTraineeData لتشمل الشهادات الخارجية
const originalLoadTraineeData = loadTraineeData;
loadTraineeData = async function() {
  await originalLoadTraineeData.call(this);
  await loadExternalCerts();
};

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
