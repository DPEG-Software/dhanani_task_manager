// Compact, Instagram-style notification center. It derives activity from the
// authoritative assignment cache and existing notification feed; it does not
// create another task/message store.
(function () {
  const READ_KEY = 'dpeg_notification_center_read_ids';
  const BASELINE_KEY = 'dpeg_notification_center_baseline_at';
  const MAX_VISIBLE = 8;
  let readIds = loadReadIds();
  let baselineAt = loadBaselineAt();
  let historyMode = false;
  let knownEventIds = null;
  let audioReady = false;
  let lastIconBadge = -1;
  const originalFaviconHref = document.querySelector('link[rel="icon"]')?.href || 'icon.svg';

  // Browsers allow sound only after a user gesture. Arm it once, then future
  // live notifications can play a short generated chime without an audio file.
  function armNotificationAudio() { audioReady = true; }
  ['pointerdown','keydown'].forEach(type => document.addEventListener(type, armNotificationAudio, { once: true, passive: true }));

  function playNotificationTone() {
    if (!audioReady || document.hidden) return;
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();
      const gain = ctx.createGain();
      gain.gain.setValueAtTime(.0001, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(.075, ctx.currentTime + .018);
      gain.gain.exponentialRampToValueAtTime(.0001, ctx.currentTime + .42);
      gain.connect(ctx.destination);
      [0, .13].forEach((delay, index) => {
        const osc = ctx.createOscillator();
        osc.type = 'sine';
        osc.frequency.value = index ? 880 : 660;
        osc.connect(gain);
        osc.start(ctx.currentTime + delay);
        osc.stop(ctx.currentTime + delay + .24);
      });
      setTimeout(() => ctx.close().catch(() => {}), 700);
    } catch (err) { console.warn('Notification tone unavailable:', err?.message || err); }
  }

  function updateAppIconBadge(count) {
    const value = Math.max(0, Number(count) || 0);
    if (value === lastIconBadge) return;
    lastIconBadge = value;
    try {
      if (value && navigator.setAppBadge) navigator.setAppBadge(value).catch(() => {});
      else if (!value && navigator.clearAppBadge) navigator.clearAppBadge().catch(() => {});
    } catch {}

    const link = document.querySelector('link[rel="icon"]');
    if (!link) return;
    if (!value) { link.href = originalFaviconHref; return; }
    const image = new Image();
    image.onload = () => {
      if (lastIconBadge !== value) return;
      const canvas = document.createElement('canvas');
      canvas.width = canvas.height = 64;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(image, 0, 0, 64, 64);
      ctx.fillStyle = '#dc2626';
      ctx.beginPath(); ctx.arc(48, 16, 16, 0, Math.PI * 2); ctx.fill();
      ctx.strokeStyle = '#fff'; ctx.lineWidth = 3; ctx.stroke();
      ctx.fillStyle = '#fff'; ctx.font = 'bold 19px Arial'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(value > 9 ? '9+' : String(value), 48, 16.5);
      link.href = canvas.toDataURL('image/png');
    };
    image.src = originalFaviconHref;
  }

  function enabled() {
    return true;
  }
  function loadReadIds() {
    try { return new Set(JSON.parse(localStorage.getItem(READ_KEY) || '[]')); }
    catch { return new Set(); }
  }
  function loadBaselineAt() {
    const stored=Number(localStorage.getItem(BASELINE_KEY)||0);
    if(stored>0)return stored;
    const now=Date.now();
    try{localStorage.setItem(BASELINE_KEY,String(now));}catch{}
    return now;
  }
  function saveReadIds() {
    try { localStorage.setItem(READ_KEY, JSON.stringify([...readIds].slice(-2000))); } catch {}
  }
  function timeOf(value) {
    const n = new Date(value || 0).getTime();
    return Number.isFinite(n) ? n : 0;
  }
  function assignmentStage(a) {
    const proof = String(a.proofStatus || '').toLowerCase();
    if (a.cancelledAt || String(a.status).toLowerCase() === 'cancelled') return 'cancelled';
    if (proof === 'approved' || String(a.status).toLowerCase() === 'done') return 'approved';
    if (proof === 'declined' || proof === 'changes_requested') return 'changes';
    if (proof === 'pending' || proof === 'submitted' || String(a.status).toLowerCase() === 'submitted') return 'proof';
    return 'assigned';
  }
  function item(id, type, title, detail, at, assignment, received) {
    return { id, type, title: title || 'Task', detail, at, assignmentId: assignment?.id || '', received };
  }
  function assignmentItems(a, received) {
    const output = [];
    const title = a.title || 'Task';
    const other = received ? (a.assignerName || a.assignerEmail || 'Someone') : (a.recipientName || a.recipientEmail || 'Someone');
    const stage = assignmentStage(a);
    if (received) output.push(item(`assigned:${a.id}:${a.createdAt || ''}`, 'assigned', title, `New task from ${other}`, a.createdAt, a, true));
    if (stage === 'proof' && !received) output.push(item(`proof:${a.id}:${a.proofSubmittedAt || a.updatedAt || ''}`, 'proof', title, `${other} submitted proof`, a.proofSubmittedAt || a.updatedAt, a, false));
    if (stage === 'changes' && received) output.push(item(`changes:${a.id}:${a.proofReviewedAt || a.updatedAt || ''}`, 'changes', title, 'Changes requested', a.proofReviewedAt || a.updatedAt, a, true));
    if (stage === 'approved' && received) output.push(item(`approved:${a.id}:${a.proofReviewedAt || a.updatedAt || ''}`, 'approved', title, 'Task approved', a.proofReviewedAt || a.updatedAt, a, true));
    if (stage === 'cancelled' && received) output.push(item(`cancelled:${a.id}:${a.cancelledAt || a.updatedAt || ''}`, 'cancelled', title, 'Task cancelled', a.cancelledAt || a.updatedAt, a, true));
    if (received && Number(a.reminderCount || 0) > 0 && a.updateAlertAt) output.push(item(`reminder:${a.id}:${a.reminderCount}`, 'reminder', title, `${a.reminderCount} reminder${Number(a.reminderCount) === 1 ? '' : 's'} received`, a.updateAlertAt, a, true));
    return output;
  }
  function followupItems(snapshot) {
    const mine = String(typeof currentUser !== 'undefined' ? currentUser?.email || '' : '').toLowerCase();
    return Object.values(window._taskFollowupState || {}).flatMap(thread => {
      const messages = Array.isArray(thread.thread) ? thread.thread : [];
      const incoming = [...messages].reverse().find(m => String(m.email || '').toLowerCase() !== mine);
      if (!incoming) return [];
      const all = [...(snapshot.assignedToMe || []), ...(snapshot.assignedByMe || [])];
      const a = all.find(row => String(row.appTaskId) === String(thread.appTaskId) && String(row.recipientEmail || '').toLowerCase() === String(thread.recipientEmail || '').toLowerCase());
      if (!a) return [];
      const received = String(a.recipientEmail || '').toLowerCase() === mine;
      return [item(`message:${incoming.id || thread.updatedAt}`, 'message', a.title || thread.taskTitle, `New message from ${incoming.name || incoming.email || 'Someone'}`, incoming.createdAt || thread.updatedAt, a, received)];
    });
  }
  function feed() {
    const snapshot = window.getTasksNotificationSnapshot?.() || { assignedToMe: [], assignedByMe: [] };
    const events = [
      ...(snapshot.assignedToMe || []).flatMap(a => assignmentItems(a, true)),
      ...(snapshot.assignedByMe || []).flatMap(a => assignmentItems(a, false)),
      ...followupItems(snapshot),
    ];
    const unique = new Map();
    events.filter(e => e.at).forEach(e => unique.set(e.id, e));
    return [...unique.values()].sort((a, b) => timeOf(b.at) - timeOf(a.at));
  }
  function icon(type) {
    return ({ assigned: '✓', message: '💬', reminder: '🔔', proof: '📎', changes: '↻', approved: '✓', cancelled: '×' })[type] || '•';
  }
  function relativeTime(value) {
    const seconds = Math.max(0, Math.round((Date.now() - timeOf(value)) / 1000));
    if (seconds < 60) return 'Now';
    if (seconds < 3600) return `${Math.floor(seconds / 60)}m`;
    if (seconds < 86400) return `${Math.floor(seconds / 3600)}h`;
    return `${Math.floor(seconds / 86400)}d`;
  }
  function render() {
    const host = document.getElementById('notification-center');
    if (!host) return;
    host.hidden = !enabled();
    if (host.hidden) return;
    const events = feed();
    const unread = events.filter(e => timeOf(e.at)>baselineAt&&!readIds.has(e.id));
    if (knownEventIds === null) {
      knownEventIds = new Set(events.map(e => e.id));
    } else {
      const hasNew = unread.some(e => !knownEventIds.has(e.id));
      events.forEach(e => knownEventIds.add(e.id));
      if (hasNew) playNotificationTone();
    }
    updateAppIconBadge(unread.length);
    const count = document.getElementById('notification-bell-count');
    if (count) { count.textContent = unread.length > 99 ? '99+' : unread.length; count.hidden = !unread.length; }
    const subtitle = document.getElementById('notification-popover-subtitle');
    if (subtitle) subtitle.textContent = historyMode ? 'Notification history' : (unread.length ? `${unread.length} unread` : "You're all caught up");
    const list = document.getElementById('notification-popover-list');
    if (!list) return;
    const visibleEvents = historyMode ? events : events.slice(0, MAX_VISIBLE);
    list.innerHTML = visibleEvents.length ? visibleEvents.map(e => `
      <button type="button" class="notification-item${readIds.has(e.id) ? '' : ' unread'}" onclick="openCenterNotification('${encodeURIComponent(e.id)}')">
        <span class="notification-item-icon type-${e.type}">${icon(e.type)}</span>
        <span class="notification-item-copy"><strong>${escapeHtml(e.title)}</strong><span>${escapeHtml(e.detail)}</span></span>
        <span class="notification-item-time">${relativeTime(e.at)}</span>
      </button>`).join('') : '<div class="notification-empty">No notifications yet</div>';
    const historyButton = document.querySelector('.notification-history-link');
    if (historyButton) historyButton.textContent = historyMode ? 'Back to recent notifications' : 'View notification history';
    window._notificationCenterFeed = events;
  }

  window.updateNotificationCenter = render;
  window.toggleNotificationCenter = function toggleNotificationCenter(event) {
    event?.stopPropagation();
    const popover = document.getElementById('notification-popover');
    const bell = document.getElementById('notification-bell');
    if (!popover) return;
    popover.hidden = !popover.hidden;
    bell?.setAttribute('aria-expanded', String(!popover.hidden));
    if (!popover.hidden) render();
  };
  window.openCenterNotification = function openCenterNotification(encodedId) {
    const id = decodeURIComponent(encodedId);
    const event = (window._notificationCenterFeed || []).find(e => e.id === id);
    if (!event) return;
    readIds.add(id); saveReadIds(); render();
    document.getElementById('notification-popover').hidden = true;
    window.openAssignmentFromNotification?.(event.assignmentId, event.received);
  };
  window.markAllCenterNotificationsRead = function markAllCenterNotificationsRead() {
    feed().forEach(e => readIds.add(e.id)); saveReadIds(); render();
  };
  window.openNotificationHistory = function openNotificationHistory() {
    historyMode = !historyMode;
    render();
  };
  document.addEventListener('click', event => {
    if (!event.target.closest('#notification-center')) {
      const popover = document.getElementById('notification-popover');
      if (popover) popover.hidden = true;
    }
  });
  window.addEventListener('load', () => setTimeout(render, 500));
})();
