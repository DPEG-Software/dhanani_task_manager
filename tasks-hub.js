(function () {
  const ASSIGNMENT_STATUSES = ['Assigned', 'Accepted', 'In Progress', 'Done'];

  let tasksTabMode = 'received'; // 'received' | 'given'
  let tasksTabCache = { assignedToMe: [], assignedByMe: [] };

  function fnBaseUrl() {
    return (localStorage.getItem('dpeg_ai_fn_url') || WORKER_URL).replace(/\/?$/, '');
  }

  function assignmentBadge(status) {
    const cls = status === 'Done' ? 'bd' : 'bp';
    return `<span class="badge ${cls}">${escapeHtml(status || 'Assigned')}</span>`;
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
          summary: (task.summary || '').slice(0, 1200),
          dept: task.dept || '',
          priority: task.priority || 'Normal',
          dueDate: task.deadline || task.date || '',
          assignerEmail: task.assignedByEmail || currentUser?.email || '',
          assignerName: task.assignedByName || currentUser?.name || '',
          recipientEmail: task.email,
          recipientName: task.person || '',
          recipientTodoListId: task.recipientTodoListId || '',
          recipientTodoTaskId: task.recipientTodoTaskId || '',
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
    const col = document.getElementById('tasks-col-person');
    if (col) col.textContent = mode === 'received' ? 'Assigned By' : 'Assigned To';
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
    tb.innerHTML = list.map(a => {
      const otherName = received ? a.assignerName : a.recipientName;
      const progressCell = received
        ? `<select class="sel-f" onchange="updateAssignmentStatus('${a.id}',this.value)">${ASSIGNMENT_STATUSES.map(s => `<option value="${s}" ${s === a.status ? 'selected' : ''}>${s}</option>`).join('')}</select>`
        : '';
      return `<tr>
        <td style="padding:10px 14px"><div style="font-size:13px;font-weight:600;color:var(--body)">${escapeHtml(a.title || '')}</div></td>
        <td style="padding:10px 10px"><div style="display:flex;align-items:center;gap:6px">${av(otherName, 22)}<span style="font-size:12px">${escapeHtml(otherName || '')}</span></div></td>
        <td style="padding:10px 8px"><span class="dept-pill"><span class="dept-dot" style="background:${dcolor(a.dept)}"></span>${escapeHtml(a.dept || '')}</span></td>
        <td style="padding:10px 8px">${assignmentBadge(a.status)}</td>
        <td style="padding:10px 8px">${pBadge(a.priority)}</td>
        <td style="padding:10px 8px">${progressCell}</td>
      </tr>`;
    }).join('');
  }

  window.updateAssignmentStatus = async function updateAssignmentStatus(id, status) {
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignment-status`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({ id, status }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const row = (tasksTabCache.assignedToMe || []).find(a => a.id === id);
      if (row) row.status = status;
      toast('Status updated');
    } catch (err) {
      console.warn('Update assignment status failed:', err.message);
      toast('Could not update status — try again');
      renderTasksTabList(); // revert the <select> to the last-known-good cached value
    }
  };
})();
