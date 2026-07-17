(function () {
  const LIVE_STAGES = ['Assigned', 'Accepted', 'In Progress', 'Submitted', 'Done'];
  const MANUAL_STATUSES = ['Assigned', 'Accepted', 'In Progress'];

  let tasksTabMode = 'received'; // 'received' | 'given'
  let tasksTabCache = { assignedToMe: [], assignedByMe: [] };
  const tasksTabOpenGroups = { received: new Set(), given: new Set() };

  function fnBaseUrl() {
    return (localStorage.getItem('dpeg_ai_fn_url') || WORKER_URL).replace(/\/?$/, '');
  }

  function groupKey(name, email) {
    return String(email || name || 'unassigned').toLowerCase();
  }

  function groupLabel(name, email) {
    return String(name || email || 'Unassigned').trim();
  }

  function groupAssignments(list, received) {
    const grouped = new Map();
    list.forEach(a => {
      const name = received ? a.assignerName : a.recipientName;
      const email = received ? a.assignerEmail : a.recipientEmail;
      const key = groupKey(name, email);
      if (!grouped.has(key)) grouped.set(key, { key, name: groupLabel(name, email), items: [] });
      grouped.get(key).items.push(a);
    });
    return [...grouped.values()].sort((a, b) => a.name.localeCompare(b.name));
  }

  function safeDomId(id) {
    return 'assign-desc-' + String(id).replace(/[^a-zA-Z0-9_-]/g, '_');
  }

  // Renders the description as readable, paragraph-preserved prose, clamped
  // with a "Show more" toggle for long forwarded-email content instead of an
  // inner scrollbar box.
  function assignmentDescription(summary, assignmentId) {
    const text = String(summary || '').trim();
    if (!text) return '';
    const clipId = safeDomId(assignmentId);
    const long = text.length > 320;
    return `<div class="assign-desc-wrap">
      <div class="assign-desc wed-sum-clip" id="${clipId}">${escapeHtml(text)}</div>
      ${long ? `<button type="button" class="wed-expand-btn" onclick="toggleAssignDescExpand('${clipId}',this)">Show more</button>` : ''}
    </div>`;
  }

  window.toggleAssignDescExpand = function toggleAssignDescExpand(clipId, btn) {
    const el = document.getElementById(clipId);
    if (!el) return;
    const expanded = el.classList.toggle('expanded');
    if (btn) btn.textContent = expanded ? 'Show less' : 'Show more';
  };

  function proofState(a) {
    return String(a?.proofStatus || 'none').toLowerCase();
  }

  // Derives the single live stage (0-4, LIVE_STAGES index) an assignment is
  // in. Proof is the source of truth once submitted — it always overrides
  // the recipient-controlled `status` field, so "Done" can only ever be
  // reached through approval, never picked directly from the dropdown.
  function assignmentStage(a) {
    const proof = proofState(a);
    if (proof === 'approved') return { index: 4, declined: false };
    if (proof === 'declined') return { index: 3, declined: true };
    if (proof === 'submitted') return { index: 3, declined: false };
    const idx = MANUAL_STATUSES.indexOf(a.status);
    return { index: idx < 0 ? 0 : idx, declined: false };
  }

  function stageLabel(a) {
    const { index, declined } = assignmentStage(a);
    return declined ? 'Declined' : LIVE_STAGES[index];
  }

  function stageSummary(items) {
    const counts = {};
    items.forEach(item => {
      const label = stageLabel(item);
      counts[label] = (counts[label] || 0) + 1;
    });
    return [...LIVE_STAGES, 'Declined']
      .filter(label => counts[label])
      .map(label => `${counts[label]} ${label}`)
      .join(' · ');
  }

  function renderStepper(a) {
    const { index, declined } = assignmentStage(a);
    const labels = ['Assigned', 'Accepted', 'In Progress', declined ? 'Declined' : 'Submitted', 'Done'];
    return `<div class="assign-stepper">${labels.map((label, i) => {
      const isDeclinedDot = declined && i === 3;
      const state = isDeclinedDot ? 'is-declined' : i < index ? 'is-complete' : i === index ? 'is-current' : '';
      const dot = `<div class="assign-step ${state}"><span class="assign-step-dot"></span><span class="assign-step-label">${label}</span></div>`;
      if (i === labels.length - 1) return dot;
      const lineComplete = !declined && i < index;
      return dot + `<span class="assign-step-line ${lineComplete ? 'is-complete' : ''}"></span>`;
    }).join('')}</div>`;
  }

  function dueDateBadge(a) {
    if (!a.dueDate) return '';
    const overdue = a.dueDate < new Date().toISOString().slice(0, 10) && stageLabel(a) !== 'Done';
    return `<span class="assign-due${overdue ? ' is-overdue' : ''}">Due ${fmtD(a.dueDate)}${overdue ? ' (overdue)' : ''}</span>`;
  }

  function assignmentActions(a, received) {
    const proof = proofState(a);
    if (received) {
      if (proof === 'none') {
        const opts = MANUAL_STATUSES.map(s => `<option value="${s}" ${s === (a.status || 'Assigned') ? 'selected' : ''}>${s}</option>`).join('');
        return `<select class="sel-f" onchange="updateAssignmentStatus('${a.id}',this.value)">${opts}</select>
          <button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Submit Proof</button>`;
      }
      if (proof === 'submitted') return `<span style="font-size:11.5px;color:var(--muted);font-weight:700">Submitted — waiting on approval</span>`;
      if (proof === 'declined') return `<span style="font-size:11.5px;color:var(--ruby);font-weight:700">Declined — check your email, then</span>
          <button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Resubmit Proof</button>`;
      if (proof === 'approved') return `<span style="font-size:11.5px;color:var(--forest);font-weight:700">✓ Approved &amp; complete</span>`;
      return '';
    }
    if (proof === 'submitted') return `<button class="btn btn-primary btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">Review Proof</button>`;
    if (proof === 'declined') return `<span style="font-size:11.5px;color:var(--ruby);font-weight:700">Declined — awaiting resubmission</span>`;
    if (proof === 'approved') return `<span style="font-size:11.5px;color:var(--forest);font-weight:700">✓ Approved &amp; complete</span>`;
    return `<span style="font-size:11.5px;color:var(--muted);font-weight:600">In progress</span>`;
  }

  function assignmentCard(a, received) {
    return `<div class="wed-card">
      <div class="wed-card-head">
        <div class="wed-card-title">${escapeHtml(a.title || '')}</div>
        <span class="dept-pill"><span class="dept-dot" style="background:${dcolor(a.dept)}"></span>${escapeHtml(a.dept || '')}</span>
        ${pBadge(a.priority)}
      </div>
      <div class="wed-card-body">
        <div class="assign-card-meta">${dueDateBadge(a)}</div>
        ${assignmentDescription(a.summary, a.id)}
        ${renderStepper(a)}
        <div class="assign-actions">${assignmentActions(a, received)}</div>
      </div>
    </div>`;
  }

  // Create/update a shared assignment record in D1 via the Worker.
  // Fire-and-forget: failures are logged only, never block To Do/OneDrive writes.
  window.recordAssignment = async function recordAssignment(task) {
    if (!task || !task.email || !task.email.includes('@dhananipeg.com')) return;
    try {
      task.assignmentId = task.assignmentId || crypto.randomUUID();
      const userToken = await getAccessToken();
      await fetch(`${fnBaseUrl()}/assignment`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({
          id: task.assignmentId,
          appTaskId: String(task.id || ''),
          title: task.title || 'Task',
          summary: (task.summary || '').slice(0, 8000),
          dept: task.dept || '',
          priority: task.priority || 'Normal',
          dueDate: task.deadline || task.date || '',
          assignerEmail: task.assignedByEmail || currentUser?.email || '',
          assignerName: task.assignedByName || currentUser?.name || '',
          recipientEmail: task.email,
          recipientName: task.person || '',
          recipientTodoListId: task.recipientTodoListId || '',
          recipientTodoTaskId: task.recipientTodoTaskId || '',
          proofInstructions: task.proofInstructions || '',
        }),
      });
    } catch (err) {
      console.warn('Assignment record (D1) failed:', err.message);
    }
  };

  window.setTasksTabMode = function setTasksTabMode(mode) {
    tasksTabMode = mode;
    document.getElementById('tasks-received-btn')?.classList.toggle('active', mode === 'received');
    document.getElementById('tasks-given-btn')?.classList.toggle('active', mode === 'given');
    const desc = document.getElementById('tasks-tab-description');
    if (desc) {
      desc.textContent = mode === 'received'
        ? 'Tasks assigned by others to you are present here.'
        : 'Tasks assigned by you to others are present here.';
    }
    renderTasksTabList();
  };

  window.renderMyTasks = async function renderMyTasks() {
    const tb = document.getElementById('tasks-tbody');
    if (!tb || !currentUser?.email) return;
    tb.innerHTML = `<div class="empty-state"><div class="es-text">Loading...</div></div>`;
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignments?email=${encodeURIComponent(currentUser.email)}`, {
        headers: { Authorization: `Bearer ${userToken}` },
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      tasksTabCache = await res.json();
    } catch (err) {
      console.warn('Load assignments failed:', err.message);
      tb.innerHTML = `<div class="empty-state"><div class="es-text">Couldn't load Tasks tab</div><div class="es-sub">Check your connection and reopen this tab to retry</div></div>`;
      return;
    }
    renderTasksTabList();
  };

  function renderTasksTabList() {
    const tb = document.getElementById('tasks-tbody');
    if (!tb) return;
    const received = tasksTabMode === 'received';
    const list = received ? (tasksTabCache.assignedToMe || []) : (tasksTabCache.assignedByMe || []);
    if (!list.length) {
      tb.innerHTML = `<div class="empty-state"><div class="es-text">No tasks yet</div></div>`;
      return;
    }
    const openGroups = tasksTabOpenGroups[tasksTabMode];
    tb.innerHTML = groupAssignments(list, received).map(group => {
      const open = openGroups.has(group.key);
      const noun = group.items.length === 1 ? 'task' : 'tasks';
      const summaryText = stageSummary(group.items);
      const safeGroupKey = escapeHtml(JSON.stringify(group.key));
      const cards = open
        ? `<div class="assign-cards">${group.items.map(a => assignmentCard(a, received)).join('')}</div>`
        : '';
      return `<div class="assign-group">
        <div class="assign-group-head" onclick="toggleTasksGroup(${safeGroupKey})">
          <span class="assign-group-toggle">${open ? '−' : '+'}</span>
          ${av(group.name, 24)}
          <span class="assign-group-name">${escapeHtml(group.name)}</span>
          <span class="assign-group-summary">${group.items.length} ${noun}${summaryText ? ` · ${escapeHtml(summaryText)}` : ''}</span>
        </div>
        ${cards}
      </div>`;
    }).join('');
  }

  window.toggleTasksGroup = function toggleTasksGroup(key) {
    const openGroups = tasksTabOpenGroups[tasksTabMode];
    if (openGroups.has(key)) openGroups.delete(key);
    else openGroups.add(key);
    renderTasksTabList();
  };

  // Opens the existing full-screen proof-submission UI (showProofUploadMode)
  // directly, without navigating away from the Tasks tab.
  window.openProofFromTasksTab = function openProofFromTasksTab(id) {
    const a = (tasksTabCache.assignedToMe || []).find(x => x.id === id);
    if (!a) return;
    showProofUploadMode({
      appTaskId: a.appTaskId || '',
      recipientEmail: currentUser?.email || '',
      assignedByName: a.assignerName || '',
      assignedByEmail: a.assignerEmail || '',
      title: a.title || '',
      proofInstructions: a.proofInstructions || '',
      proofShareUrl: '',
      todoListId: '',
      todoTaskId: '',
    });
  };

  window.openProofReviewFromTasksTab = async function openProofReviewFromTasksTab(id) {
    const a = (tasksTabCache.assignedByMe || []).find(x => x.id === id);
    if (!a) return;
    if (typeof checkAndLoadProofNotifications === 'function') {
      await checkAndLoadProofNotifications().catch(err => console.warn('Proof notification refresh failed:', err.message));
    }
    if (typeof nav === 'function') nav('notifications');
    toast('Open Notifications to review the submitted proof');
  };

  window.updateTasksTabProofState = function updateTasksTabProofState(appTaskId, proofStatus) {
    const key = String(appTaskId || '');
    if (!key) return;
    const now = new Date().toISOString();
    const changedRows = [
      ...(tasksTabCache.assignedToMe || []).filter(a => String(a.appTaskId || '') === key),
      ...(tasksTabCache.assignedByMe || []).filter(a => String(a.appTaskId || '') === key),
    ];
    changedRows.forEach(row => {
      row.proofStatus = proofStatus;
      if (proofStatus === 'submitted') {
        row.proofSubmittedAt = now;
        row.proofReviewedAt = null;
      }
      if (proofStatus === 'approved' || proofStatus === 'declined') row.proofReviewedAt = now;
      if (proofStatus === 'approved') row.status = 'Done';
    });
    renderTasksTabList();
  };

  window.updateAssignmentStatus = async function updateAssignmentStatus(id, status) {
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignment-status`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({ id, status }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const changedRows = [
        ...(tasksTabCache.assignedToMe || []).filter(a => a.id === id),
        ...(tasksTabCache.assignedByMe || []).filter(a => a.id === id),
      ];
      changedRows.forEach(row => { row.status = status; });
      renderTasksTabList();
      toast('Progress updated');
    } catch (err) {
      console.warn('Update assignment status failed:', err.message);
      toast('Could not update progress — try again');
      renderTasksTabList(); // revert the <select> to the last-known-good cached value
    }
  };
})();
