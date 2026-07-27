(function () {
  const LIVE_STAGES = ['Assigned', 'In Progress', 'Submitted', 'Done'];
  const MANUAL_STATUSES = ['Assigned', 'In Progress'];

  let tasksTabMode = 'received'; // 'received' | 'given'
  let tasksTabCache = { assignedToMe: [], assignedByMe: [] };
  const tasksTabOpenGroups = { received: new Set(), given: new Set() };

  // "New" tracking for the red count badge next to each group's name. Keyed by
  // id + current stage (not just id) so a task that was already seen still
  // re-alerts when its stage changes later (e.g. a delegator gets alerted
  // again once a recipient submits proof, not just when the task was first
  // assigned). Persisted to localStorage so it survives reloads.
  const SEEN_STORAGE_KEY = 'dpeg_seen_assignment_stages';
  function loadSeenStages() {
    try { return new Set(JSON.parse(localStorage.getItem(SEEN_STORAGE_KEY) || '[]')); }
    catch { return new Set(); }
  }
  function saveSeenStages(set) {
    try { localStorage.setItem(SEEN_STORAGE_KEY, JSON.stringify([...set].slice(-500))); } catch {}
  }
  let seenAssignmentStages = loadSeenStages();
  function assignmentSeenKey(a) { return `${a.id}::${stageLabel(a)}`; }

  // Follow-up "seen" tracking is separate from the stage-badge Set above,
  // since it needs to remember *how many* thread messages were already seen
  // (to compute an unread count), not just a single seen/unseen flag.
  const FOLLOWUP_SEEN_STORAGE_KEY = 'dpeg_followup_seen_lengths';
  function loadFollowupSeenLengths() {
    try { return JSON.parse(localStorage.getItem(FOLLOWUP_SEEN_STORAGE_KEY) || '{}'); }
    catch { return {}; }
  }
  function saveFollowupSeenLengths(map) {
    try { localStorage.setItem(FOLLOWUP_SEEN_STORAGE_KEY, JSON.stringify(map)); } catch {}
  }
  let followupSeenLengths = loadFollowupSeenLengths();

  function fnBaseUrl() {
    return (localStorage.getItem('dpeg_ai_fn_url') || WORKER_URL).replace(/\/?$/, '');
  }

  function groupKey(name, email) {
    return String(email || name || 'unassigned').toLowerCase();
  }

  function groupLabel(name, email) {
    return String(name || email || 'Unassigned').trim();
  }

  // Completed tasks sink to the bottom of their group; everything else stays
  // newest-assigned-first, so a task that just finished doesn't linger mixed
  // in among active ones, and the most recent active task is what you see first.
  function sortAssignmentItems(a, b) {
    const aDone = stageLabel(a) === 'Done';
    const bDone = stageLabel(b) === 'Done';
    if (aDone !== bDone) return aDone ? 1 : -1;
    return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
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
    const groups = [...grouped.values()].sort((a, b) => a.name.localeCompare(b.name));
    groups.forEach(g => g.items.sort(sortAssignmentItems));
    return groups;
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
    // Always show the toggle, even for a short one-line description — kept
    // unconditional rather than measuring overflow, since that was flaky
    // across zoom levels/card widths and still missed cases in practice.
    return `<div class="assign-desc-wrap">
      <div class="assign-desc wed-sum-clip" id="${clipId}">${escapeHtml(text)}</div>
      <button type="button" class="wed-expand-btn" onclick="toggleAssignDescExpand('${clipId}',this)">Show more</button>
    </div>`;
  }

  window.toggleAssignDescExpand = function toggleAssignDescExpand(clipId, btn) {
    const el = document.getElementById(clipId);
    if (!el) return;
    const expanded = el.classList.toggle('expanded');
    if (btn) btn.textContent = expanded ? 'Show less' : 'Show more';
  };

  // The "Show more" button always shows (per request), but the bottom fade
  // overlay should only appear when a description is actually clamped —
  // otherwise a short one-line description gets a fade smeared across it
  // for no reason. Toggles a class the CSS keys off; doesn't touch the button.
  function syncAssignDescClamped(container) {
    requestAnimationFrame(() => {
      container.querySelectorAll('.assign-desc').forEach(clip => {
        clip.classList.toggle('is-clamped', clip.scrollHeight > clip.clientHeight + 1);
      });
    });
  }

  function proofState(a) {
    return String(a?.proofStatus || 'none').toLowerCase();
  }

  // Derives the single live stage (0-3, LIVE_STAGES index) an assignment is
  // in. Proof is the source of truth once submitted — it always overrides
  // the recipient-controlled `status` field, so "Done" can only ever be
  // reached through approval, never picked directly from the dropdown.
  function assignmentStage(a) {
    const proof = proofState(a);
    if (proof === 'approved') return { index: 3, declined: false };
    if (proof === 'declined') return { index: 2, declined: true };
    if (proof === 'submitted') return { index: 2, declined: false };
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
    const labels = ['Assigned', 'In Progress', declined ? 'Declined' : 'Submitted', 'Done'];
    return `<div class="assign-stepper">${labels.map((label, i) => {
      const isDeclinedDot = declined && i === 2;
      // The terminal stage (Done) is a finished state, not an in-flight one —
      // render it solid/complete (and specially highlighted) rather than the
      // "current" in-progress ring.
      const isFinal = i === labels.length - 1;
      const isDone = isFinal && i === index && !isDeclinedDot;
      const state = isDeclinedDot ? 'is-declined' : i < index ? 'is-complete' : i === index ? (isFinal ? 'is-complete' : 'is-current') : '';
      const dot = `<div class="assign-step ${state}${isDone ? ' is-done' : ''}"><span class="assign-step-dot"></span><span class="assign-step-label">${label}</span></div>`;
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

  // Follow-up thread state is populated by index.html's checkAndLoadProofNotifications
  // poll (window._taskFollowupState), keyed by appTaskId+recipientEmail so both the
  // Received and Delegated view of the same task share one conversation.
  function followupThreadState(a) {
    const key = `${a.appTaskId}::${String(a.recipientEmail || '').toLowerCase()}`;
    return window._taskFollowupState ? window._taskFollowupState[key] : null;
  }

  // Count of thread messages from the other party since this user last
  // opened (or sent into) this task's follow-up thread. Using a count
  // rather than a boolean means two separate follow-up messages read as
  // "2", not just "unread".
  function followupUnreadCount(a, received) {
    const state = followupThreadState(a);
    if (!state || !Array.isArray(state.thread) || !state.thread.length) return 0;
    const myRole = received ? 'assignee' : 'assignor';
    const seenLen = followupSeenLengths[a.id] || 0;
    return state.thread.slice(seenLen).filter(m => m && m.by !== myRole).length;
  }

  function followupButton(a, received) {
    const count = followupUnreadCount(a, received);
    const badge = count > 0 ? `<span class="assign-followup-count">${count > 9 ? '9+' : count}</span>` : '';
    return `<button class="btn btn-ghost btn-sm assign-followup-btn" onclick="openTaskFollowup('${a.id}',${received})">Follow up${badge}</button>`;
  }

  function assignmentActions(a, received) {
    const proof = proofState(a);
    const followBtn = followupButton(a, received);
    let content;
    if (received) {
      if (proof === 'none') {
        const opts = MANUAL_STATUSES.map(s => `<option value="${s}" ${s === (a.status || 'Assigned') ? 'selected' : ''}>${s}</option>`).join('');
        content = `<select class="sel-f" onchange="updateAssignmentStatus('${a.id}',this.value)">${opts}</select>
          <button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Submit Proof</button>`;
      } else if (proof === 'submitted') {
        content = `<span style="font-size:11.5px;color:var(--muted);font-weight:700">Submitted — waiting on approval</span>
          <button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Update Proof</button>`;
      } else if (proof === 'declined') {
        content = `<span style="font-size:11.5px;color:var(--ruby);font-weight:700">Declined — check your email, then</span>
          <button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Resubmit Proof</button>`;
      } else if (proof === 'approved') {
        content = `<span style="font-size:11.5px;color:var(--forest);font-weight:700">✓ Approved &amp; complete</span>`;
      } else {
        content = '';
      }
    } else if (proof === 'submitted') {
      content = `<button class="btn btn-primary btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">Review Proof</button>`;
    } else if (proof === 'declined') {
      content = `<span style="font-size:11.5px;color:var(--ruby);font-weight:700">Declined — awaiting resubmission</span>`;
    } else if (proof === 'approved') {
      content = `<span style="font-size:11.5px;color:var(--forest);font-weight:700">✓ Approved &amp; complete</span>`;
    } else {
      content = `<span style="font-size:11.5px;color:var(--muted);font-weight:600">In progress</span>`;
    }
    return `${content}${followBtn}`;
  }

  function assignmentCard(a, received) {
    const hasFollowup = followupUnreadCount(a, received) > 0;
    return `<div class="wed-card${hasFollowup ? ' has-followup' : ''}">
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

  function assignmentsSignature(cache) {
    const sig = list => (list || [])
      .map(a => [a.id, a.status, a.proofStatus, a.summary, a.dueDate, a.title, a.dept, a.priority].join('|'))
      .join(';');
    return `${sig(cache?.assignedToMe)}::${sig(cache?.assignedByMe)}`;
  }

  // silent=true is used by the background poll: fetches quietly and only
  // touches the DOM if something actually changed, so a card the user has
  // expanded doesn't flash/collapse on every refresh cycle.
  window.renderMyTasks = async function renderMyTasks(silent) {
    const tb = document.getElementById('tasks-tbody');
    if (!tb || !currentUser?.email) return;
    if (!silent) tb.innerHTML = `<div class="empty-state"><div class="es-text">Loading...</div></div>`;
    let nextCache;
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignments?email=${encodeURIComponent(currentUser.email)}`, {
        headers: { Authorization: `Bearer ${userToken}` },
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      nextCache = await res.json();
    } catch (err) {
      console.warn('Load assignments failed:', err.message);
      if (!silent) {
        tb.innerHTML = `<div class="empty-state"><div class="es-text">Couldn't load Tasks tab</div><div class="es-sub">Check your connection and reopen this tab to retry</div></div>`;
      }
      return;
    }
    if (silent && assignmentsSignature(nextCache) === assignmentsSignature(tasksTabCache)) return;
    tasksTabCache = nextCache;
    renderTasksTabList();
  };

  // Alerts the receiver of a new assignment (and the assignor of a new
  // submission, since that's also an unseen stage change) without requiring
  // the Tasks tab to be open: a badge on the sidebar "Tasks" nav item, and
  // one on each of the Received/Delegated toggle buttons.
  function updateTasksNavBadges() {
    const toMe = tasksTabCache.assignedToMe || [];
    const byMe = tasksTabCache.assignedByMe || [];
    const unseenCount = list => list.filter(a => !seenAssignmentStages.has(assignmentSeenKey(a))).length;
    const receivedCount = unseenCount(toMe);
    const givenCount = unseenCount(byMe);
    const setBadge = (id, count) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.textContent = count > 99 ? '99+' : count;
      el.style.display = count > 0 ? '' : 'none';
    };
    setBadge('nb-tasks', receivedCount + givenCount);
    setBadge('tasks-received-badge', receivedCount);
    setBadge('tasks-given-badge', givenCount);
  }

  function renderTasksTabList() {
    const tb = document.getElementById('tasks-tbody');
    if (!tb) return;
    updateTasksNavBadges();
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
      const newCount = group.items.filter(a => !seenAssignmentStages.has(assignmentSeenKey(a))).length;
      const newBadge = newCount > 0 ? `<span class="assign-new-badge">${newCount > 9 ? '9+' : newCount}</span>` : '';
      const followupTotal = group.items.reduce((sum, a) => sum + followupUnreadCount(a, received), 0);
      const followupGroupBadge = followupTotal > 0
        ? `<span class="assign-followup-group-badge">+${followupTotal > 9 ? '9+' : followupTotal} follow-up${followupTotal > 1 ? 's' : ''}</span>`
        : '';
      const cards = open
        ? `<div class="assign-cards">${group.items.map(a => assignmentCard(a, received)).join('')}</div>`
        : '';
      return `<div class="assign-group">
        <div class="assign-group-head" onclick="toggleTasksGroup(${safeGroupKey})">
          <span class="assign-group-toggle">${open ? '−' : '+'}</span>
          <span class="assign-avatar-wrap">${av(group.name, 24)}${newBadge}</span>
          <span class="assign-group-name">${escapeHtml(group.name)}</span>
          ${followupGroupBadge}
          <span class="assign-group-summary">${group.items.length} ${noun}${summaryText ? ` · ${escapeHtml(summaryText)}` : ''}</span>
        </div>
        ${cards}
      </div>`;
    }).join('');
    syncAssignDescClamped(tb);
  }

  window.toggleTasksGroup = function toggleTasksGroup(key) {
    const openGroups = tasksTabOpenGroups[tasksTabMode];
    if (openGroups.has(key)) {
      openGroups.delete(key);
    } else {
      openGroups.add(key);
      // Expanding a group is what clears its "new" badge — mark everything
      // currently in it as seen at its current stage.
      const received = tasksTabMode === 'received';
      const list = received ? (tasksTabCache.assignedToMe || []) : (tasksTabCache.assignedByMe || []);
      const group = groupAssignments(list, received).find(g => g.key === key);
      if (group) {
        group.items.forEach(a => seenAssignmentStages.add(assignmentSeenKey(a)));
        saveSeenStages(seenAssignmentStages);
      }
    }
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

  // Opens the generic task follow-up modal (index.html: showTaskFollowupModal),
  // available on both Received and Delegated cards regardless of proof status.
  window.openTaskFollowup = function openTaskFollowup(id, received) {
    const list = received ? (tasksTabCache.assignedToMe || []) : (tasksTabCache.assignedByMe || []);
    const a = list.find(x => x.id === id);
    if (!a || typeof window.showTaskFollowupModal !== 'function') return;
    window.showTaskFollowupModal({
      assignmentId: a.id,
      appTaskId: a.appTaskId || '',
      title: a.title || '',
      assignerEmail: a.assignerEmail || '',
      assignerName: a.assignerName || '',
      recipientEmail: a.recipientEmail || '',
      recipientName: a.recipientName || '',
      role: received ? 'assignee' : 'assignor',
    });
  };

  // Called by showTaskFollowupModal once the thread has loaded, so opening
  // the modal clears the unread count the same way expanding a group clears
  // its "new" badge. Stores the thread length seen so far (not just a flag)
  // so a later re-open can compute exactly how many new messages arrived.
  window.markTaskFollowupSeen = function markTaskFollowupSeen(assignmentId, threadLen) {
    if (!assignmentId) return;
    followupSeenLengths[assignmentId] = threadLen;
    saveFollowupSeenLengths(followupSeenLengths);
    renderTasksTabList();
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
