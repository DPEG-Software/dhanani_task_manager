(function () {
  const ASSIGNMENT_STATUSES = ['Assigned', 'Accepted', 'In Progress', 'Done'];

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

  function assignmentDescription(summary) {
    const text = String(summary || '').trim();
    if (!text) return '';
    return `<div style="margin-top:7px;padding:8px 10px;background:var(--bg);border:1px solid var(--border);border-radius:6px;font-size:11.5px;line-height:1.55;color:var(--sub);white-space:pre-wrap;max-height:180px;overflow:auto">${escapeHtml(text)}</div>`;
  }

  function statusSummary(items) {
    const counts = {};
    items.forEach(item => {
      const status = item.status || 'Assigned';
      counts[status] = (counts[status] || 0) + 1;
    });
    return ASSIGNMENT_STATUSES
      .filter(status => counts[status])
      .map(status => `${counts[status]} ${status}`)
      .join(' · ');
  }

  function proofState(a) {
    return String(a?.proofStatus || 'none').toLowerCase();
  }

  function proofBadge(status) {
    const s = String(status || 'none').toLowerCase();
    const styles = {
      submitted: ['Proof submitted', '#fffbeb', '#f59e0b', '#92400e'],
      approved: ['Proof approved', '#f0fdf4', '#86efac', '#166534'],
      declined: ['Proof declined', '#fff1f2', '#fca5a5', '#b91c1c'],
    };
    const cfg = styles[s];
    if (!cfg) return '';
    return `<span style="display:inline-flex;align-items:center;white-space:nowrap;padding:2px 8px;background:${cfg[1]};border:1px solid ${cfg[2]};border-radius:10px;font-size:10.5px;font-weight:800;color:${cfg[3]}">${cfg[0]}</span>`;
  }

  function proofSummary(items) {
    const counts = {};
    items.forEach(item => {
      const status = proofState(item);
      if (status === 'none') return;
      counts[status] = (counts[status] || 0) + 1;
    });
    return ['submitted', 'approved', 'declined']
      .filter(status => counts[status])
      .map(status => `${counts[status]} proof ${status}`)
      .join(' · ');
  }

  function proofCell(a, received) {
    const status = proofState(a);
    const badge = proofBadge(status);
    if (received) {
      if (status === 'submitted') {
        return `<div style="display:flex;flex-direction:column;align-items:flex-start;gap:5px">${badge}<span style="font-size:11px;color:var(--muted)">Under review</span></div>`;
      }
      if (status === 'approved') return badge;
      if (status === 'declined') {
        return `<div style="display:flex;flex-direction:column;align-items:flex-start;gap:6px">${badge}<button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Resubmit Proof</button></div>`;
      }
      return `<button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Submit Proof</button>`;
    }
    if (!badge) return '';
    const reviewButton = status === 'submitted'
      ? `<button class="btn btn-ghost btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">Review</button>`
      : '';
    return `<div style="display:flex;flex-direction:column;align-items:flex-start;gap:6px">${badge}${reviewButton}</div>`;
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
          summary: (task.summary || '').slice(0, 4000),
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
    tb.innerHTML = `<tr><td colspan="6"><div class="empty-state"><div class="es-text">Loading...</div></div></td></tr>`;
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignments?email=${encodeURIComponent(currentUser.email)}`, {
        headers: { Authorization: `Bearer ${userToken}` },
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      tasksTabCache = await res.json();
    } catch (err) {
      console.warn('Load assignments failed:', err.message);
      tb.innerHTML = `<tr><td colspan="6"><div class="empty-state"><div class="es-text">Couldn't load Tasks tab</div><div class="es-sub">Check your connection and reopen this tab to retry</div></div></td></tr>`;
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
      tb.innerHTML = `<tr><td colspan="6"><div class="empty-state"><div class="es-text">No tasks yet</div></div></td></tr>`;
      return;
    }
    const openGroups = tasksTabOpenGroups[tasksTabMode];
    tb.innerHTML = groupAssignments(list, received).map(group => {
      const open = openGroups.has(group.key);
      const noun = group.items.length === 1 ? 'task' : 'tasks';
      const progressSummary = statusSummary(group.items);
      const proofGroupSummary = proofSummary(group.items);
      const rows = open ? group.items.map(a => {
        const progressCell = received
          ? `<select class="sel-f" onchange="updateAssignmentStatus('${a.id}',this.value)">${ASSIGNMENT_STATUSES.map(s => `<option value="${s}" ${s === a.status ? 'selected' : ''}>${s}</option>`).join('')}</select>`
          : `<span style="font-size:12px;color:var(--muted)">${escapeHtml(a.status || 'Assigned')}</span>`;
        return `<tr>
          <td style="padding:10px 14px"><div style="font-size:13px;font-weight:600;color:var(--body)">${escapeHtml(a.title || '')}</div>${assignmentDescription(a.summary)}</td>
          <td style="padding:10px 8px"><span class="dept-pill"><span class="dept-dot" style="background:${dcolor(a.dept)}"></span>${escapeHtml(a.dept || '')}</span></td>
          <td style="padding:10px 8px">${fmtD(a.dueDate)}</td>
          <td style="padding:10px 8px">${pBadge(a.priority)}</td>
          <td style="padding:10px 8px">${progressCell}</td>
          <td style="padding:10px 8px">${proofCell(a, received)}</td>
        </tr>`;
      }).join('') : '';
      const summaryText = [progressSummary, proofGroupSummary].filter(Boolean).join(' · ');
      const safeGroupKey = escapeHtml(JSON.stringify(group.key));
      return `<tr onclick="toggleTasksGroup(${safeGroupKey})" style="background:var(--sage3);cursor:pointer">
        <td colspan="6" style="padding:10px 14px">
          <div style="display:flex;align-items:center;gap:10px;min-width:0">
            <span style="font-size:13px;color:var(--muted);width:14px;display:inline-flex;justify-content:center">${open ? '-' : '+'}</span>
            ${av(group.name, 24)}
            <div style="min-width:0;flex:1">
              <div style="font-size:13px;font-weight:800;color:var(--body)">${escapeHtml(group.name)}</div>
            </div>
            <div style="font-size:11px;color:var(--sub);font-weight:700;white-space:nowrap">${group.items.length} ${noun}${summaryText ? ` · ${escapeHtml(summaryText)}` : ''}</div>
          </div>
        </td>
      </tr>${rows}`;
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
