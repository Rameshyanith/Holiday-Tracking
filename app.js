'use strict';

// ═══════════════════════════════════════════════════════════════
// UK BANK HOLIDAYS (England, 2025–2027)
// ═══════════════════════════════════════════════════════════════
const UK_BANK_HOLIDAYS = new Set([
  '2025-01-01', '2025-04-18', '2025-04-21', '2025-05-05',
  '2025-05-26', '2025-08-25', '2025-12-25', '2025-12-26',
  '2026-01-01', '2026-04-03', '2026-04-06', '2026-05-04',
  '2026-05-25', '2026-08-31', '2026-12-25', '2026-12-28',
  '2027-01-01', '2027-03-26', '2027-03-29', '2027-05-03',
  '2027-05-31', '2027-08-30', '2027-12-27', '2027-12-28',
]);

const BH_LABELS = {
  '2026-01-01': "New Year's Day",
  '2026-04-03': 'Good Friday',
  '2026-04-06': 'Easter Monday',
  '2026-05-04': 'Early May Bank Holiday',
  '2026-05-25': 'Spring Bank Holiday',
  '2026-08-31': 'Summer Bank Holiday',
  '2026-12-25': 'Christmas Day',
  '2026-12-28': 'Boxing Day (substitute)',
};

const USER_COLORS = {
  alice: '#6366f1', bob: '#ec4899', carol: '#22c55e', dave: '#f59e0b', admin: '#06b6d4',
};

// ═══════════════════════════════════════════════════════════════
// STORE
// ═══════════════════════════════════════════════════════════════
const Store = {
  init() {
    if (!localStorage.getItem('hts_v1')) {
      localStorage.setItem('hts_users', JSON.stringify([
        { id: 'alice', name: 'Alice', password: 'password123', role: 'user', allowance: 28 },
        { id: 'bob', name: 'Bob', password: 'password123', role: 'user', allowance: 28 },
        { id: 'carol', name: 'Carol', password: 'password123', role: 'user', allowance: 28 },
        { id: 'dave', name: 'Dave', password: 'password123', role: 'user', allowance: 28 },
        { id: 'admin', name: 'Admin', password: 'admin123', role: 'admin', allowance: 28 },
      ]));
      localStorage.setItem('hts_holidays', JSON.stringify([]));
      localStorage.setItem('hts_v1', '1');
    }
  },
  getUsers() { return JSON.parse(localStorage.getItem('hts_users') || '[]'); },
  saveUsers(u) { localStorage.setItem('hts_users', JSON.stringify(u)); },
  getHolidays() { return JSON.parse(localStorage.getItem('hts_holidays') || '[]'); },
  saveHolidays(h) { localStorage.setItem('hts_holidays', JSON.stringify(h)); },
  getCurrentUser() {
    const id = sessionStorage.getItem('hts_session');
    return id ? this.getUsers().find(u => u.id === id) || null : null;
  },
  setCurrentUser(id) { sessionStorage.setItem('hts_session', id); },
  clearSession() { sessionStorage.removeItem('hts_session'); },
};

// ═══════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════
function uid() { return Math.random().toString(36).slice(2, 9) + Date.now().toString(36); }
function toStr(d) { return d.toISOString().split('T')[0]; }
function cloneDate(s) { return new Date(s + 'T00:00:00'); }

function formatDate(s) {
  return cloneDate(s).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function workingDays(startStr, endStr) {
  let n = 0;
  const cur = cloneDate(startStr), end = cloneDate(endStr);
  while (cur <= end) {
    const dow = cur.getDay();
    if (dow !== 0 && dow !== 6 && !UK_BANK_HOLIDAYS.has(toStr(cur))) n++;
    cur.setDate(cur.getDate() + 1);
  }
  return n;
}

function getUserStats(userId) {
  const yr = new Date().getFullYear().toString();
  const user = Store.getUsers().find(u => u.id === userId);
  const mine = Store.getHolidays().filter(h => h.userId === userId && h.startDate.startsWith(yr));
  const approved = mine.filter(h => h.status === 'approved').reduce((s, h) => s + h.days, 0);
  const pending = mine.filter(h => h.status === 'pending').reduce((s, h) => s + h.days, 0);
  // Bank holidays in this year count as free days (part of the 28-day statutory entitlement)
  const bhCount = [...UK_BANK_HOLIDAYS].filter(d => d.startsWith(yr)).length;
  const bookable = user.allowance - bhCount;        // days that can actually be booked
  const remaining = bookable - approved - pending;
  return { allowance: user.allowance, bookable, approved, pending, remaining, bhCount };
}

function overlaps(s1, e1, s2, e2) { return s1 <= e2 && e1 >= s2; }

function greeting() {
  const h = new Date().getHours();
  if (h < 12) return 'Good morning';
  if (h < 17) return 'Good afternoon';
  return 'Good evening';
}

// Build a pre-filled Outlook Web calendar deep-link URL
function outlookCalendarUrl(h, userName) {
  // Outlook all-day events use an exclusive end date, so add 1 day
  const endExclusive = new Date(h.endDate + 'T00:00:00');
  endExclusive.setDate(endExclusive.getDate() + 1);
  const endStr = toStr(endExclusive);
  const subject = encodeURIComponent(`Annual Leave – ${userName}`);
  const body = encodeURIComponent(
    `Holiday booked via Holiday Tracking System.\n` +
    `Duration: ${h.days} working day${h.days !== 1 ? 's' : ''}` +
    (h.reason ? `\nReason: ${h.reason}` : '')
  );
  return `https://outlook.office.com/calendar/action/compose?subject=${subject}&startdt=${h.startDate}&enddt=${endStr}&allday=true&body=${body}`;
}

// ═══════════════════════════════════════════════════════════════
// TOAST
// ═══════════════════════════════════════════════════════════════
function toast(msg, type = 'info') {
  const icons = { success: 'fa-circle-check', error: 'fa-circle-xmark', info: 'fa-circle-info', warning: 'fa-triangle-exclamation' };
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i><span>${msg}</span>`;
  document.getElementById('toast-container').appendChild(el);
  requestAnimationFrame(() => { requestAnimationFrame(() => el.classList.add('show')); });
  setTimeout(() => { el.classList.remove('show'); setTimeout(() => el.remove(), 350); }, 3500);
}

// ═══════════════════════════════════════════════════════════════
// ROUTER
// ═══════════════════════════════════════════════════════════════
const Router = {
  navigate(view) {
    const user = Store.getCurrentUser();
    if (view !== 'login' && !user) { this.navigate('login'); return; }
    if (view === 'admin' && user?.role !== 'admin') { this.navigate('dashboard'); return; }

    // Login vs App
    const loginEl = document.getElementById('view-login');
    const appEl = document.getElementById('app');
    if (view === 'login') {
      loginEl.classList.remove('hidden'); appEl.classList.add('hidden');
    } else {
      loginEl.classList.add('hidden'); appEl.classList.remove('hidden');
    }

    // Views
    document.querySelectorAll('.app-view').forEach(v => v.classList.remove('active'));
    const target = document.getElementById(`view-${view}`);
    if (target) target.classList.add('active');

    // Nav
    document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
    const navEl = document.querySelector(`.nav-link[data-view="${view}"]`);
    if (navEl) navEl.classList.add('active');

    updateSidebar();
    Views[view]?.render();
  }
};

// ═══════════════════════════════════════════════════════════════
// SIDEBAR / UI
// ═══════════════════════════════════════════════════════════════
function updateSidebar() {
  const user = Store.getCurrentUser();
  if (!user) return;

  const av = document.getElementById('nav-user-avatar');
  av.textContent = user.name[0];
  av.style.background = USER_COLORS[user.id] || '#6366f1';
  document.getElementById('nav-user-name').textContent = user.name;
  document.getElementById('nav-user-role').textContent = user.role === 'admin' ? 'Administrator' : 'Employee';

  // Admin link
  const adminLink = document.getElementById('nav-admin-link');
  adminLink.classList.toggle('hidden', user.role !== 'admin');

  // Pending badge for admin
  const pending = Store.getHolidays().filter(h => h.status === 'pending');
  const adminBadge = document.getElementById('nav-badge-admin');
  if (user.role === 'admin' && pending.length > 0) {
    adminBadge.textContent = pending.length;
    adminBadge.classList.remove('hidden');
  } else {
    adminBadge.classList.add('hidden');
  }

  // My holidays pending badge
  const myPending = pending.filter(h => h.userId === user.id).length;
  const myBadge = document.getElementById('nav-badge-pending');
  if (myPending > 0) {
    myBadge.textContent = myPending;
    myBadge.classList.remove('hidden');
  } else {
    myBadge.classList.add('hidden');
  }
}

// ═══════════════════════════════════════════════════════════════
// VIEWS
// ═══════════════════════════════════════════════════════════════
const Views = {

  // ─── LOGIN ───────────────────────────────────────────────
  login: {
    render() {
      const sel = document.getElementById('login-name');
      sel.innerHTML = '<option value="">Select your name…</option>';
      Store.getUsers().forEach(u => {
        const opt = document.createElement('option');
        opt.value = u.name; opt.textContent = u.name;
        sel.appendChild(opt);
      });
      document.getElementById('login-error').classList.add('hidden');
      document.getElementById('login-password').value = '';
    }
  },

  // ─── DASHBOARD ───────────────────────────────────────────
  dashboard: {
    calDate: new Date(),
    render() {
      const user = Store.getCurrentUser();
      const stats = getUserStats(user.id);

      document.getElementById('dash-greeting').textContent = `${greeting()}, ${user.name}!`;
      document.getElementById('dash-year').textContent = new Date().getFullYear();
      document.getElementById('stat-allowance').textContent = stats.allowance;
      document.getElementById('stat-approved').textContent = stats.approved;
      document.getElementById('stat-pending').textContent = stats.pending;
      document.getElementById('stat-remaining').textContent = Math.max(0, stats.remaining);

      const pct = stats.bookable > 0 ? Math.min(100, Math.round(((stats.approved + stats.pending) / stats.bookable) * 100)) : 0;
      document.getElementById('days-progress').style.width = pct + '%';
      document.getElementById('days-progress-pct').textContent = pct + '%';
      document.getElementById('days-progress-label').textContent =
        `${stats.approved} approved + ${stats.pending} pending of ${stats.bookable} bookable days`;

      this.renderCalendar();
      this.renderUpcoming();
    },

    renderCalendar() {
      const d = this.calDate, year = d.getFullYear(), month = d.getMonth();
      document.getElementById('cal-month-label').textContent =
        d.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });

      const firstDow = (new Date(year, month, 1).getDay() + 6) % 7; // Mon=0
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      const todayStr = toStr(new Date());

      // Build day→entries map for this month
      const dayMap = {};
      const yrMo = `${year}-${String(month + 1).padStart(2, '0')}`;
      Store.getHolidays()
        .filter(h => h.status !== 'rejected' && h.startDate.slice(0, 7) <= yrMo && h.endDate.slice(0, 7) >= yrMo)
        .forEach(h => {
          const cur = cloneDate(h.startDate), end = cloneDate(h.endDate);
          while (cur <= end) {
            if (cur.getMonth() === month && cur.getFullYear() === year) {
              const k = toStr(cur);
              if (!dayMap[k]) dayMap[k] = [];
              dayMap[k].push(h);
            }
            cur.setDate(cur.getDate() + 1);
          }
        });

      const grid = document.getElementById('cal-grid');
      grid.innerHTML = '';
      ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].forEach(h => {
        const el = document.createElement('div'); el.className = 'cal-header'; el.textContent = h; grid.appendChild(el);
      });
      for (let i = 0; i < firstDow; i++) {
        const el = document.createElement('div'); el.className = 'cal-day empty'; grid.appendChild(el);
      }
      for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dow = cloneDate(dateStr).getDay();
        const isBH = UK_BANK_HOLIDAYS.has(dateStr);
        const el = document.createElement('div');
        el.className = 'cal-day' + (dow === 0 || dow === 6 ? ' weekend' : '') + (isBH ? ' bank-holiday' : '') + (dateStr === todayStr ? ' today' : '');
        if (isBH) el.title = `Bank Holiday${BH_LABELS[dateStr] ? ' — ' + BH_LABELS[dateStr] : ''}`;

        const numEl = document.createElement('span'); numEl.className = 'cal-day-num'; numEl.textContent = day; el.appendChild(numEl);

        const entries = dayMap[dateStr];
        if (entries?.length) {
          const row = document.createElement('div'); row.className = 'cal-dots-row';
          entries.forEach(h => {
            const dot = document.createElement('div'); dot.className = 'cal-dot';
            dot.style.background = USER_COLORS[h.userId] || '#6366f1';
            dot.style.opacity = h.status === 'pending' ? '0.55' : '1';
            dot.title = `${Store.getUsers().find(u => u.id === h.userId)?.name || h.userId} (${h.status})`;
            row.appendChild(dot);
          });
          el.appendChild(row);
        }
        grid.appendChild(el);
      }

      // Legend
      const legend = document.getElementById('cal-legend');
      const regularUsers = Store.getUsers().filter(u => u.role === 'user');
      legend.innerHTML = regularUsers.map(u =>
        `<div class="legend-item"><span class="legend-dot" style="background:${USER_COLORS[u.id]}"></span>${u.name}</div>`
      ).join('') + `<div class="legend-item"><span class="legend-dot bh-dot"></span>Bank Holiday</div>`;
    },

    renderUpcoming() {
      const user = Store.getCurrentUser(), today = toStr(new Date());
      const upcoming = Store.getHolidays()
        .filter(h => h.userId === user.id && h.endDate >= today && h.status !== 'rejected')
        .sort((a, b) => a.startDate.localeCompare(b.startDate)).slice(0, 4);
      const el = document.getElementById('upcoming-list');
      if (!upcoming.length) {
        el.innerHTML = `<div class="empty-state"><i class="fas fa-umbrella-beach"></i><p>No upcoming holidays booked</p></div>`;
        return;
      }
      el.innerHTML = upcoming.map(h => {
        const outlookBtn = h.status === 'approved'
          ? `<a href="${outlookCalendarUrl(h, user.name)}" target="_blank" rel="noopener" class="outlook-icon-btn" title="Add to Outlook Calendar"><i class="fas fa-calendar-plus"></i></a>`
          : '';
        return `
        <div class="upcoming-item">
          <div class="upcoming-dates"><i class="fas fa-calendar-alt"></i>
            <span>${formatDate(h.startDate)} – ${formatDate(h.endDate)}</span>
          </div>
          <div class="upcoming-meta">
            <span class="badge badge-${h.status}">${h.status}</span>
            <span class="upcoming-days">${h.days} day${h.days !== 1 ? 's' : ''}</span>
            ${outlookBtn}
          </div>
        </div>`;
      }).join('');
    }
  },

  // ─── BOOK ─────────────────────────────────────────────────
  book: {
    render() {
      document.getElementById('book-form').reset();
      document.getElementById('book-days-preview').classList.add('hidden');
      const today = toStr(new Date());
      document.getElementById('book-start').min = today;
      document.getElementById('book-end').min = today;

      // Populate 2026 bank holidays list
      const ul = document.getElementById('bh-list-2026');
      if (ul) ul.innerHTML = Object.entries(BH_LABELS).map(([d, n]) =>
        `<li><strong>${formatDate(d)}</strong> — ${n}</li>`).join('');
    }
  },

  // ─── MY HOLIDAYS ──────────────────────────────────────────
  'my-holidays': {
    render() {
      const user = Store.getCurrentUser();
      const list = Store.getHolidays()
        .filter(h => h.userId === user.id)
        .sort((a, b) => b.requestedAt.localeCompare(a.requestedAt));
      const tbody = document.getElementById('my-holidays-tbody');
      if (!list.length) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-state-td">
          <div class="empty-state"><i class="fas fa-calendar-times"></i><p>No holiday requests yet — go book one!</p></div>
        </td></tr>`; return;
      }
      tbody.innerHTML = list.map(h => {
        let action = '—';
        if (h.status === 'pending') {
          action = `<button class="btn btn-sm btn-danger" onclick="cancelHoliday('${h.id}')"><i class="fas fa-xmark"></i> Cancel</button>`;
        } else if (h.status === 'approved') {
          const url = outlookCalendarUrl(h, user.name);
          action = `<a href="${url}" target="_blank" rel="noopener" class="btn btn-sm btn-outlook" title="Add to Outlook Calendar"><i class="fas fa-calendar-plus"></i> Add to Outlook</a>`;
        }
        return `
        <tr>
          <td>${formatDate(h.startDate)}</td>
          <td>${formatDate(h.endDate)}</td>
          <td><span class="days-badge">${h.days}</span></td>
          <td style="max-width:180px;color:var(--tx3);font-size:.78rem">${h.reason || '—'}</td>
          <td><span class="badge badge-${h.status}">${h.status}</span></td>
          <td>${action}</td>
        </tr>`;
      }).join('');
    }
  },

  // ─── ADMIN ────────────────────────────────────────────────
  admin: {
    render() {
      this.renderTeamGrid();
      this.renderPending();
      this.renderAll();
    },
    renderTeamGrid() {
      const users = Store.getUsers().filter(u => u.role === 'user');
      const el = document.getElementById('admin-team-grid');
      el.innerHTML = users.map(u => {
        const s = getUserStats(u.id);
        return `<div class="admin-team-card">
          <div class="team-avatar" style="background:${USER_COLORS[u.id]}">${u.name[0]}</div>
          <div>
            <div class="team-name">${u.name}</div>
            <div class="team-stats">${s.remaining} days left · ${s.approved} approved</div>
          </div>
        </div>`;
      }).join('');
    },
    renderPending() {
      const pending = Store.getHolidays().filter(h => h.status === 'pending')
        .sort((a, b) => a.requestedAt.localeCompare(b.requestedAt));
      const users = Store.getUsers();
      const badge = document.getElementById('pending-count-badge');
      badge.textContent = pending.length ? `${pending.length} pending` : 'None';
      badge.className = pending.length ? 'badge badge-pending' : 'badge badge-approved';

      const el = document.getElementById('admin-pending-list');
      if (!pending.length) {
        el.innerHTML = `<div class="empty-state" style="padding:1.5rem 1.4rem">
          <i class="fas fa-circle-check"></i><p>No pending requests — all clear!</p>
        </div>`; return;
      }
      el.innerHTML = pending.map(h => {
        const u = users.find(x => x.id === h.userId);
        return `<div class="admin-request-card" id="req-${h.id}">
          <div class="req-user">
            <div class="req-avatar" style="background:${USER_COLORS[h.userId]}">${u?.name[0] || '?'}</div>
            <div>
              <div class="req-name">${u?.name || h.userId}</div>
              <div class="req-date">Submitted ${formatDate(h.requestedAt.split('T')[0])}</div>
            </div>
          </div>
          <div class="req-info">
            <div><i class="fas fa-calendar"></i> ${formatDate(h.startDate)} → ${formatDate(h.endDate)}</div>
            <div><i class="fas fa-sun"></i> ${h.days} working day${h.days !== 1 ? 's' : ''}</div>
            ${h.reason ? `<div class="req-reason"><i class="fas fa-comment"></i> ${h.reason}</div>` : ''}
          </div>
          <div class="req-actions">
            <button class="btn btn-success" onclick="approveHol('${h.id}')"><i class="fas fa-check"></i> Approve</button>
            <button class="btn btn-danger"  onclick="rejectHol('${h.id}')"><i class="fas fa-xmark"></i> Reject</button>
          </div>
        </div>`;
      }).join('');
    },
    renderAll() {
      const filterUser = document.getElementById('admin-filter-user')?.value || '';
      const filterStatus = document.getElementById('admin-filter-status')?.value || '';

      // Populate user filter
      const sel = document.getElementById('admin-filter-user');
      if (sel && sel.options.length <= 1) {
        Store.getUsers().filter(u => u.role === 'user').forEach(u => {
          const o = document.createElement('option'); o.value = u.id; o.textContent = u.name; sel.appendChild(o);
        });
      }

      let list = Store.getHolidays().sort((a, b) => b.requestedAt.localeCompare(a.requestedAt));
      if (filterUser) list = list.filter(h => h.userId === filterUser);
      if (filterStatus) list = list.filter(h => h.status === filterStatus);
      const users = Store.getUsers();
      const tbody = document.getElementById('admin-all-tbody');
      if (!list.length) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-state-td">
          <div class="empty-state"><i class="fas fa-database"></i><p>No records match this filter</p></div>
        </td></tr>`; return;
      }
      tbody.innerHTML = list.map(h => {
        const u = users.find(x => x.id === h.userId);
        return `<tr>
          <td><div class="user-cell"><span class="user-dot" style="background:${USER_COLORS[h.userId]}"></span>${u?.name || h.userId}</div></td>
          <td>${formatDate(h.startDate)}</td>
          <td>${formatDate(h.endDate)}</td>
          <td><span class="days-badge">${h.days}</span></td>
          <td><span class="badge badge-${h.status}">${h.status}</span></td>
          <td>${h.status === 'pending'
            ? `<button class="btn btn-sm btn-success" onclick="approveHol('${h.id}')"><i class="fas fa-check"></i></button>
               <button class="btn btn-sm btn-danger"  onclick="rejectHol('${h.id}')"><i class="fas fa-xmark"></i></button>`
            : '—'}
          </td>
        </tr>`;
      }).join('');
    }
  },

  // ─── SETTINGS ─────────────────────────────────────────────
  settings: {
    render() {
      const user = Store.getCurrentUser();
      const stats = getUserStats(user.id);
      const av = document.getElementById('settings-avatar');
      av.textContent = user.name[0];
      av.style.background = USER_COLORS[user.id] || '#6366f1';
      document.getElementById('settings-username').textContent = user.name;
      document.getElementById('settings-role-badge').textContent = user.role === 'admin' ? 'Administrator' : 'Employee';
      document.getElementById('settings-stat-remaining').textContent = Math.max(0, stats.remaining);
      document.getElementById('settings-stat-approved').textContent = stats.approved;
      document.getElementById('settings-form').reset();
      // Pre-fill new-username field with current name
      document.getElementById('settings-new-username').value = user.name;
      document.getElementById('settings-username-pw').value = '';
    }
  }
};

// ═══════════════════════════════════════════════════════════════
// GLOBAL ACTIONS (called from inline onclick)
// ═══════════════════════════════════════════════════════════════
function approveHol(id) {
  const hols = Store.getHolidays(), h = hols.find(x => x.id === id);
  if (!h) return;
  h.status = 'approved';
  Store.saveHolidays(hols);
  toast(`Approved holiday for ${Store.getUsers().find(u => u.id === h.userId)?.name}`, 'success');
  Views.admin.render(); updateSidebar();
}
function rejectHol(id) {
  const hols = Store.getHolidays(), h = hols.find(x => x.id === id);
  if (!h) return;
  h.status = 'rejected';
  Store.saveHolidays(hols);
  toast('Holiday request rejected', 'error');
  Views.admin.render(); updateSidebar();
}
function cancelHoliday(id) {
  const hols = Store.getHolidays(), idx = hols.findIndex(x => x.id === id);
  if (idx < 0) return;
  hols.splice(idx, 1);
  Store.saveHolidays(hols);
  toast('Holiday request cancelled', 'info');
  Views['my-holidays'].render(); updateSidebar();
}

// ═══════════════════════════════════════════════════════════════
// EVENT HANDLERS
// ═══════════════════════════════════════════════════════════════
function bindEvents() {

  // ── Login ──
  document.getElementById('login-form').addEventListener('submit', e => {
    e.preventDefault();
    const name = document.getElementById('login-name').value;
    const pw = document.getElementById('login-password').value;
    const user = Store.getUsers().find(u => u.name.toLowerCase() === name.toLowerCase() && u.password === pw);
    const errEl = document.getElementById('login-error');
    if (user) {
      errEl.classList.add('hidden');
      Store.setCurrentUser(user.id);
      Router.navigate('dashboard');
      toast(`Welcome back, ${user.name}!`, 'success');
    } else {
      errEl.classList.remove('hidden');
    }
  });

  // Password toggle on login
  document.getElementById('toggle-pw').addEventListener('click', () => {
    const inp = document.getElementById('login-password');
    const ico = document.querySelector('#toggle-pw i');
    inp.type = inp.type === 'password' ? 'text' : 'password';
    ico.className = inp.type === 'password' ? 'fas fa-eye' : 'fas fa-eye-slash';
  });

  // ── Logout ──
  document.getElementById('btn-logout').addEventListener('click', () => {
    Store.clearSession();
    Router.navigate('login');
    toast('Signed out successfully', 'info');
  });

  // ── Nav links ──
  document.querySelectorAll('.nav-link').forEach(link => {
    link.addEventListener('click', e => {
      e.preventDefault();
      Router.navigate(link.dataset.view);
      document.getElementById('sidebar').classList.remove('open');
    });
  });

  // ── Sidebar toggle (mobile) ──
  document.getElementById('sidebar-toggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('open');
  });

  // ── Calendar prev/next ──
  document.getElementById('cal-prev').addEventListener('click', () => {
    Views.dashboard.calDate.setMonth(Views.dashboard.calDate.getMonth() - 1);
    Views.dashboard.renderCalendar();
  });
  document.getElementById('cal-next').addEventListener('click', () => {
    Views.dashboard.calDate.setMonth(Views.dashboard.calDate.getMonth() + 1);
    Views.dashboard.renderCalendar();
  });

  // ── Book form: live preview ──
  function updatePreview() {
    const s = document.getElementById('book-start').value;
    const e = document.getElementById('book-end').value;
    const prev = document.getElementById('book-days-preview');
    if (s && e && e >= s) {
      const n = workingDays(s, e);
      prev.textContent = n > 0
        ? `${n} working day${n !== 1 ? 's' : ''} (weekends & bank holidays excluded)`
        : 'No working days in this range';
      prev.classList.remove('hidden');
      prev.style.color = n === 0 ? 'var(--red)' : '';
    } else {
      prev.classList.add('hidden');
    }
    if (s) document.getElementById('book-end').min = s;
  }
  document.getElementById('book-start').addEventListener('change', updatePreview);
  document.getElementById('book-end').addEventListener('change', updatePreview);

  // ── Book form: submit ──
  document.getElementById('book-form').addEventListener('submit', e => {
    e.preventDefault();
    const start = document.getElementById('book-start').value;
    const end = document.getElementById('book-end').value;
    const reason = document.getElementById('book-reason').value.trim();

    if (!start || !end || end < start) { toast('Please select valid start and end dates', 'error'); return; }

    const days = workingDays(start, end);
    if (days === 0) { toast('No working days fall in this range', 'warning'); return; }

    const user = Store.getCurrentUser();
    const stats = getUserStats(user.id);
    if (days > stats.remaining) { toast(`Not enough days remaining. You have ${Math.max(0, stats.remaining)} bookable day${stats.remaining !== 1 ? 's' : ''} left`, 'error'); return; }

    const myHols = Store.getHolidays().filter(h => h.userId === user.id && h.status !== 'rejected');
    if (myHols.some(h => overlaps(start, end, h.startDate, h.endDate))) {
      toast('These dates overlap with an existing request', 'error'); return;
    }

    const hols = Store.getHolidays();
    hols.push({ id: uid(), userId: user.id, startDate: start, endDate: end, days, reason, status: 'pending', requestedAt: new Date().toISOString() });
    Store.saveHolidays(hols);
    toast(`Holiday request submitted for ${days} day${days !== 1 ? 's' : ''}!`, 'success');
    Router.navigate('my-holidays');
  });

  // ── Admin filters ──
  document.getElementById('admin-filter-user')?.addEventListener('change', () => Views.admin.renderAll());
  document.getElementById('admin-filter-status')?.addEventListener('change', () => Views.admin.renderAll());

  // ── Settings: change password ──
  document.getElementById('settings-form').addEventListener('submit', e => {
    e.preventDefault();
    const current = document.getElementById('settings-current-pw').value;
    const newPw = document.getElementById('settings-new-pw').value;
    const confirm = document.getElementById('settings-confirm-pw').value;

    if (newPw !== confirm) { toast('New passwords do not match', 'error'); return; }
    if (newPw.length < 6) { toast('Password must be at least 6 characters', 'error'); return; }

    const user = Store.getCurrentUser();
    const users = Store.getUsers();
    const u = users.find(x => x.id === user.id);
    if (!u || u.password !== current) { toast('Current password is incorrect', 'error'); return; }

    u.password = newPw;
    Store.saveUsers(users);
    toast('Password updated successfully!', 'success');
    document.getElementById('settings-form').reset();
  });

  // ── Settings: change username ──
  document.getElementById('settings-username-form').addEventListener('submit', e => {
    e.preventDefault();
    const newName = document.getElementById('settings-new-username').value.trim();
    const pw = document.getElementById('settings-username-pw').value;

    if (!newName) { toast('Username cannot be empty', 'error'); return; }
    if (newName.length < 2) { toast('Username must be at least 2 characters', 'error'); return; }

    const user = Store.getCurrentUser();
    const users = Store.getUsers();
    const u = users.find(x => x.id === user.id);
    if (!u || u.password !== pw) { toast('Password is incorrect', 'error'); return; }

    // Check for duplicate name (case-insensitive, ignore self)
    const duplicate = users.some(x => x.id !== user.id && x.name.toLowerCase() === newName.toLowerCase());
    if (duplicate) { toast('That username is already taken', 'error'); return; }

    u.name = newName;
    Store.saveUsers(users);
    toast('Username updated successfully!', 'success');
    // Refresh profile card & sidebar
    Views.settings.render();
    updateSidebar();
  });
}

// ═══════════════════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  Store.init();
  bindEvents();

  // Start with login view hidden to avoid flash
  document.getElementById('view-login').classList.remove('hidden');
  const user = Store.getCurrentUser();
  if (user) {
    Router.navigate('dashboard');
  } else {
    Router.navigate('login');
  }
});
