(function () {
  const LIVE_STAGES = ['Assigned', 'In Progress', 'Submitted', 'Done'];
  const MANUAL_STATUSES = ['Assigned', 'In Progress'];

  let tasksTabMode = 'received'; // 'received' | 'given' | 'department' | 'history'
  let tasksHistoryFilter = 'completed'; // 'completed' | 'cancelled'
  let tasksTabCache = { assignedToMe: [], assignedByMe: [], overseenByMe: [] };
  const tasksTabOpenGroups = { received: new Set(), given: new Set(), department: new Set(), history: new Set() };
  const tasksTabVisibleCounts = { received: new Map(), given: new Map(), department: new Map(), history: new Map() };
  const tasksHistoryDirections = new Map(); // person key -> 'to' | 'by'
  let expandedAssignmentId = null;
  const TASKS_PAGE_SIZE = 10;

  // Proof notifications created from older/manual assignments do not always
  // contain the assignor email. Let the notification loader verify ownership
  // against the authoritative D1 Delegated list instead of dropping them.
  window.isDelegatedTaskProof = function isDelegatedTaskProof(appTaskId, recipientEmail) {
    const taskKey=String(appTaskId||'');
    const recipient=String(recipientEmail||'').toLowerCase();
    return (tasksTabCache.assignedByMe||[]).some(a=>
      String(a.appTaskId||'')===taskKey&&
      (!recipient||String(a.recipientEmail||'').toLowerCase()===recipient)
    );
  };

  window.findDelegatedAssignmentByAppTaskId = function findDelegatedAssignmentByAppTaskId(appTaskId) {
    const key=String(appTaskId||'');
    return (tasksTabCache.assignedByMe||[]).find(a=>String(a.appTaskId||'')===key)||null;
  };

  // "New" tracking for the red count badge next to each group's name.
  // Assigned and In Progress are one assignment event: changing your own
  // progress must not manufacture another New badge. Proof decisions remain
  // separate events and keep their own indicators.
  const SEEN_STORAGE_KEY = 'dpeg_seen_assignment_stages';
  function loadSeenStages() {
    try { return new Set(JSON.parse(localStorage.getItem(SEEN_STORAGE_KEY) || '[]')); }
    catch { return new Set(); }
  }
  function saveSeenStages(set) {
    try { localStorage.setItem(SEEN_STORAGE_KEY, JSON.stringify([...set].slice(-500))); } catch {}
  }
  let seenAssignmentStages = loadSeenStages();
  function assignmentSeenKey(a) {
    const stage=stageLabel(a);
    const seenStage=stage==='In Progress'?'Assigned':stage;
    const eventTime=stage==='Submitted'?a.proofSubmittedAt||'':stage==='Changes Requested'?a.proofReviewedAt||'':'';
    return `${a.id}::${seenStage}::${eventTime}`;
  }

  function isNewAssignment(a) {
    return stageLabel(a)==='Assigned'&&!seenAssignmentStages.has(assignmentSeenKey(a));
  }

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

  // Person-row indicators are cleared by expanding that person, while the
  // task-level Follow Up count and reminder total remain available.
  const GROUP_NOTICE_STORAGE_KEY = 'dpeg_task_group_notice_totals';
  function loadGroupNoticeTotals() {
    try { return JSON.parse(localStorage.getItem(GROUP_NOTICE_STORAGE_KEY) || '{}'); }
    catch { return {}; }
  }
  function saveGroupNoticeTotals() {
    try { localStorage.setItem(GROUP_NOTICE_STORAGE_KEY, JSON.stringify(groupNoticeTotals)); } catch {}
  }
  let groupNoticeTotals = loadGroupNoticeTotals();

  const TASK_REMINDER_SEEN_STORAGE_KEY = 'dpeg_task_reminder_seen_counts';
  function loadTaskReminderSeenCounts() {
    try { return JSON.parse(localStorage.getItem(TASK_REMINDER_SEEN_STORAGE_KEY) || '{}'); }
    catch { return {}; }
  }
  function saveTaskReminderSeenCounts() {
    try { localStorage.setItem(TASK_REMINDER_SEEN_STORAGE_KEY, JSON.stringify(taskReminderSeenCounts)); } catch {}
  }
  let taskReminderSeenCounts = loadTaskReminderSeenCounts();

  function hasNewReminder(a, received) {
    if (!received) return false;
    const localSeen=Number(taskReminderSeenCounts[a.id]||0);
    // A null shared value means this row predates central read receipts. Do
    // not resurrect its historical total unless a live reminder is present.
    if(a.recipientReminderSeenCount==null&&!a.updateAlertAt&&!localSeen)return false;
    const sharedSeen=a.recipientReminderSeenCount==null?0:Number(a.recipientReminderSeenCount||0);
    return Math.max(0,Number(a.reminderCount||0)) > Math.max(localSeen,sharedSeen);
  }

  function groupNoticeKey(mode, key) {
    return `${mode}::${key}`;
  }

  function incomingFollowupTotal(items, received) {
    const myRole = received ? 'assignee' : 'assignor';
    return items.reduce((sum, a) => {
      const thread = followupThreadState(a)?.thread || [];
      return sum + thread.filter(message => message && message.by !== myRole).length;
    }, 0);
  }

  function groupTransientNotices(mode, group) {
    if (mode === 'history') return { followups: 0, reminders: 0 };
    const noticeKey = groupNoticeKey(mode,group.key);
    const seenNotice = groupNoticeTotals[noticeKey];
    const totalIncoming = incomingFollowupTotal(group.items,group.received);
    const totalReminders = mode === 'received'
      ? group.items.reduce((sum,a)=>sum+Math.max(0,Number(a.reminderCount||0)),0)
      : 0;
    const followupDelta=Math.max(0,seenNotice
        ? totalIncoming-Number(seenNotice.followups||0)
        : group.items.reduce((sum,a)=>sum+followupUnreadCount(a,group.received),0));
    const reminderDelta=Math.max(0,seenNotice
        ? totalReminders-Number(seenNotice.reminders||0)
        : group.items.filter(a=>hasNewReminder(a,mode==='received')).length);
    return {followups:followupDelta?1:0,reminders:reminderDelta?1:0};
  }

  function fnBaseUrl() {
    return (localStorage.getItem('dpeg_ai_fn_url') || WORKER_URL).replace(/\/?$/, '');
  }

  async function persistAssignmentSeen(a,{threadLen=0,reminderCount=0}={}){
    if(!a?.id)return;
    try{
      const userToken=await getAccessToken();
      await fetch(`${fnBaseUrl()}/assignment-seen`,{
        method:'POST',headers:{'Content-Type':'application/json',Authorization:`Bearer ${userToken}`},
        body:JSON.stringify({id:a.id,threadLen,reminderCount}),
      });
    }catch(err){console.warn('Save assignment seen state failed:',err.message);}
  }

  function groupKey(name, email) {
    return String(email || name || 'unassigned').toLowerCase();
  }

  function groupLabel(name, email) {
    return String(name || email || 'Unassigned').trim();
  }

  // Most-recent-activity signal used to order both individual cards and the
  // person-groups they sit in — a task (and the person it's grouped under)
  // should float back toward the top the moment anything happens on it
  // (status change, proof event, cancellation, or a new follow-up message),
  // not just sit wherever it landed when first assigned.
  function assignmentRecency(a) {
    const followupTs = followupThreadState(a)?.updatedAt;
    const times = [a.updatedAt, a.cancelledAt, a.proofSubmittedAt, a.proofReviewedAt, a.updateAlertAt, a.createdAt, followupTs]
      .map(t => (t ? new Date(t).getTime() : NaN))
      .filter(Number.isFinite);
    return times.length ? Math.max(...times) : 0;
  }

  // Completed tasks sink to the bottom of their group; everything else stays
  // most-recently-active-first, so a task that just finished doesn't linger
  // mixed in among active ones, and whatever just got touched (assigned,
  // updated, or replied to) is what you see first.
  function isAssignmentOverdue(a) {
    return !!a.dueDate && a.dueDate < new Date().toISOString().slice(0,10) && !['Done','Cancelled'].includes(stageLabel(a));
  }

  function sortAssignmentItems(a, b) {
    const aDone = stageLabel(a) === 'Done';
    const bDone = stageLabel(b) === 'Done';
    if (aDone !== bDone) return aDone ? 1 : -1;
    return assignmentRecency(b) - assignmentRecency(a);
  }

  // History is all Done/Cancelled already, so sort by when it was resolved
  // (approved or cancelled, falling back to any later update, then creation)
  // instead of when it was originally assigned.
  function sortHistoryItems(a, b) {
    const resolvedAt = x => x.proofReviewedAt || x.cancelledAt || x.updatedAt || x.createdAt || 0;
    return new Date(resolvedAt(b)) - new Date(resolvedAt(a));
  }

  // `received` is either a fixed boolean (Received/Delegated tabs, one
  // direction for the whole list) or a per-item lookup fn (History, which
  // merges both directions). Groups are keyed with a direction prefix only
  // in the per-item case, so a person who both assigned you something and
  // received something from you doesn't collapse into a single group.
  function groupAssignments(list, received, sortFn = sortAssignmentItems) {
    const perItem = typeof received === 'function';
    const grouped = new Map();
    list.forEach(a => {
      const isReceived = perItem ? received(a) : received;
      const name = isReceived ? a.assignerName : a.recipientName;
      const email = isReceived ? a.assignerEmail : a.recipientEmail;
      // History supplies a per-item direction, but both directions belong
      // under one person. Direction is selected inside the expanded group.
      const key = groupKey(name, email);
      if (!grouped.has(key)) grouped.set(key, { key, name: groupLabel(name, email), received: isReceived, items: [] });
      grouped.get(key).items.push(a);
    });
    const groups = [...grouped.values()];
    groups.forEach(g => g.items.sort((a,b)=>sortFn(a,b,g.received)));
    // Whoever has the most recently active task floats to the top of the
    // name list too, instead of a static alphabetical order that never
    // reflects what just happened.
    groups.sort((a, b) => assignmentRecency(b.items[0])-assignmentRecency(a.items[0]));
    return groups;
  }

  // Shared by renderTasksTabList and toggleTasksGroup so both derive the
  // exact same list/grouping for a given tab mode.
  function tasksTabModeSource(mode) {
    if (mode === 'history') {
      const wantedStage = tasksHistoryFilter === 'cancelled' ? 'Cancelled' : 'Done';
      const isPast = a => stageLabel(a) === wantedStage;
      const toMe = (tasksTabCache.assignedToMe || []).filter(isPast).map(a => ({ ...a, _received: true }));
      const byMe = (tasksTabCache.assignedByMe || []).filter(isPast).map(a => ({ ...a, _received: false }));
      return { list: [...toMe, ...byMe], received: a => a._received, sortFn: sortHistoryItems };
    }
    if (mode === 'department') {
      const list = (tasksTabCache.overseenByMe || []).filter(a => stageLabel(a) !== 'Done' && stageLabel(a) !== 'Cancelled');
      return { list, received: false, principal: true, sortFn: sortAssignmentItems };
    }
    const received = mode === 'received';
    const list = (received ? (tasksTabCache.assignedToMe || []) : (tasksTabCache.assignedByMe || []))
      .filter(a => stageLabel(a) !== 'Done' && stageLabel(a) !== 'Cancelled');
    return { list, received, principal: false, sortFn: sortAssignmentItems };
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

  function isCancelled(a) {
    return a?.status === 'Cancelled';
  }

  function awaitingApproval(a) {
    return !isCancelled(a) && (proofState(a) === 'submitted' || !!window.hasPendingTaskProofReview?.(a?.appTaskId));
  }

  // Derives the single live stage (0-3, LIVE_STAGES index) an assignment is
  // in. Proof is the source of truth once submitted — it always overrides
  // the recipient-controlled `status` field, so "Done" can only ever be
  // reached through approval, never picked directly from the dropdown.
  // Cancelled is checked first since it's assigner-set independent of proof
  // state — a task can be cancelled even mid-review, with a proof already
  // sitting as 'submitted' underneath it.
  function assignmentStage(a) {
    if (isCancelled(a)) return { index: -1, declined: false, cancelled: true };
    const proof = proofState(a);
    if (proof === 'approved') return { index: 3, declined: false };
    if (proof === 'declined') return { index: 2, declined: true };
    if (proof === 'submitted') return { index: 2, declined: false };
    const idx = MANUAL_STATUSES.indexOf(a.status);
    return { index: idx < 0 ? 0 : idx, declined: false };
  }

  function stageLabel(a) {
    const { index, declined, cancelled } = assignmentStage(a);
    return cancelled ? 'Cancelled' : declined ? 'Changes Requested' : LIVE_STAGES[index];
  }

  function stageSummary(items) {
    const counts = {};
    items.forEach(item => {
      const label = stageLabel(item);
      counts[label] = (counts[label] || 0) + 1;
    });
    return [...LIVE_STAGES, 'Changes Requested', 'Cancelled']
      .filter(label => counts[label])
      .map(label => `${counts[label]} ${label}`)
      .join(' · ');
  }

  function friendlyGroupSummary(items, received) {
    const parts=[`${items.length} task${items.length===1?'':'s'}`];
    return parts.join(' · ');
  }

  function renderStepper(a) {
    const { index, declined, cancelled } = assignmentStage(a);
    if (cancelled) {
      return `<div class="assign-cancelled-banner">✕ Cancelled${a.cancelReason ? `: ${escapeHtml(a.cancelReason)}` : ''}</div>`;
    }
    const labels = ['Assigned', 'In Progress', declined ? 'Changes Requested' : 'Submitted', 'Done'];
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
    const stage = stageLabel(a);
    const overdue = a.dueDate < new Date().toISOString().slice(0, 10) && stage !== 'Done' && stage !== 'Cancelled';
    return `<span class="assign-due${overdue ? ' is-overdue' : ''}">Due ${fmtD(a.dueDate)}${overdue ? ' (overdue)' : ''}</span>`;
  }

  // Follow-up thread state is populated by index.html's checkAndLoadProofNotifications
  // poll (window._taskFollowupState), keyed by appTaskId+recipientEmail so both the
  // Received and Delegated view of the same task share one conversation.
  function followupThreadState(a) {
    const key = `${a.appTaskId}::${String(a.recipientEmail || '').toLowerCase()}`;
    return window._taskFollowupState ? window._taskFollowupState[key] : null;
  }

  // Seen-length is keyed by assignment id *and* role, not just assignment id.
  // On a self-assigned task (assigner === recipient) the Received and
  // Delegated views point at the exact same D1 row/id — without the role in
  // the key, sending a follow-up from Delegated would mark it "seen" for
  // the Received view too, since they'd share one storage slot.
  function followupSeenKey(a, received) {
    return `${a.id}::${received ? 'assignee' : 'assignor'}`;
  }

  // Count of thread messages from the other party since this user last
  // opened (or sent into) this task's follow-up thread. Using a count
  // rather than a boolean means two separate follow-up messages read as
  // "2", not just "unread".
  function followupUnreadCount(a, received) {
    const state = followupThreadState(a);
    if (!state || !Array.isArray(state.thread) || !state.thread.length) return 0;
    const myRole = received ? 'assignee' : 'assignor';
    const sharedSeen=received?a.recipientMessageSeenCount:a.assignerMessageSeenCount;
    const seenLen = Math.max(Number(sharedSeen||0),Number(followupSeenLengths[followupSeenKey(a, received)] || 0));
    return state.thread.slice(seenLen).filter(m => m && m.by !== myRole).length;
  }

  function followupButton(a, received) {
    const count = followupUnreadCount(a, received);
    const badge = count > 0 ? '<span class="assign-followup-count">1</span>' : '';
    return `<button class="btn btn-ghost btn-sm assign-followup-btn" onclick="openTaskFollowup('${a.id}',${received})">Messages${badge}</button>`;
  }

  // Standalone one-click "update required" nudge — separate from the
  // follow-up thread entirely. Bell button lives only on Delegated cards
  // (the assignor sends it); the fixed label it produces is read on
  // Received cards/groups. Cleared server-side once the recipient updates
  // status or submits proof (see handleAssignmentStatus/updateAssignmentProofState).
  const ALERT_LABEL = 'Reminder';

  function alertBellButton(a, received) {
    if (received) return '';
    const count = Math.max(0, Number(a.reminderCount || 0));
    const label = count ? `Reminder ${count} sent` : 'Remind';
    return `<button type="button" class="btn btn-ghost btn-sm assign-alert-btn" title="Send another reminder" onclick="sendUpdateAlert('${a.id}')">🔔 ${label}</button>`;
  }

  function alertLabel(a, received) {
    if (!received || !a.updateAlertAt) return '';
    // Auto-clears once the recipient submits/updates proof or the assigner
    // resolves it (approve/decline) — but neither happens if an alert lands
    // while a proof is already sitting in review, so give the recipient an
    // explicit way to dismiss it themselves rather than being stuck.
    return `<span class="assign-alert-label">${ALERT_LABEL}<button type="button" class="assign-alert-dismiss" title="Dismiss alert" onclick="event.stopPropagation();dismissUpdateAlert('${a.id}')">&times;</button></span>`;
  }

  function moreMenuButton(a) {
    return `<button type="button" class="btn btn-ghost btn-sm assign-cancel-btn" title="Cancel task" onclick="cancelAssignmentDirect('${a.id}')">Cancel</button>`;
  }

  function assignmentActions(a, received, principal=false) {
    const proof = proofState(a);
    const cancelled = isCancelled(a);
    // A task that's Done — or Cancelled — has nothing left to follow up or
    // alert about; both are terminal states the same way, they just got
    // there through different doors. Once a task lands in either, it sinks
    // into the completed/"history" part of its group (see sortAssignmentItems
    // and tasksTabModeSource) and these controls retire along with it.
    const isTerminal = proof === 'approved' || cancelled;
    if (principal) {
      const approvalOwner=escapeHtml(a.assignerName||a.assignerEmail||'the assigner');
      const approvalText=awaitingApproval(a)
        ?`Proof awaiting approval from ${approvalOwner}`
        :proof==='approved'
        ?`✓ Approved by ${approvalOwner}`
        :`Approval owner: ${approvalOwner}`;
      return `<span style="font-size:11.5px;color:var(--muted);font-weight:700">${approvalText}</span><span class="assign-actions-trailing">${isTerminal?'':`<button class="btn btn-ghost btn-sm assign-followup-btn" onclick="openTaskFollowup('${a.id}','principal')">Messages</button>`}</span>`;
    }
    const followBtn = isTerminal ? '' : followupButton(a, received);
    const bellBtn = (isTerminal || awaitingApproval(a)) ? '' : alertBellButton(a, received);
    // Cancelling is an assigner-only action, and only makes sense while a
    // task is still live — nothing left to call off once it's Done or
    // already Cancelled.
    const moreBtn = (!received && !isTerminal) ? moreMenuButton(a) : '';
    let content;
    if (cancelled) {
      content = received
        ? `<span style="font-size:11.5px;color:var(--ruby);font-weight:700">Cancelled by ${escapeHtml(a.assignerName || a.assignerEmail || 'the assignor')}</span>`
        : `<span style="font-size:11.5px;color:var(--ruby);font-weight:700">✕ Cancelled</span>`;
    } else if (received) {
      if (proof === 'none') {
        const opts = MANUAL_STATUSES.map(s => `<option value="${s}" ${s === (a.status || 'Assigned') ? 'selected' : ''}>${s}</option>`).join('');
        content = `<select class="sel-f" onchange="updateAssignmentStatus('${a.id}',this.value)">${opts}</select>
          <button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Submit Proof</button>`;
      } else if (proof === 'submitted') {
        content = `<span style="font-size:11.5px;color:var(--muted);font-weight:700">Submitted — waiting on approval</span>`;
      } else if (proof === 'declined') {
        content = `<button class="btn btn-ghost btn-sm" onclick="openProofFromTasksTab('${a.id}')">Resubmit Proof</button>`;
      } else if (proof === 'approved') {
        content = `<span style="font-size:11.5px;color:var(--forest);font-weight:700">✓ Approved &amp; complete</span>`;
      } else {
        content = '';
      }
    } else if (awaitingApproval(a)) {
      content = `<button class="btn btn-primary btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">View Proof</button>`;
    } else if (proof === 'declined') {
      // Keep rejected submissions available to the assigner while the
      // assignee works on a correction. A later pending resubmission enters
      // the awaitingApproval branch above and changes this to View Proof.
      content = `<button class="btn btn-ghost btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">Proofs</button>`;
    } else if (proof === 'approved') {
      content = `<span style="font-size:11.5px;color:var(--forest);font-weight:700">✓ Approved &amp; complete</span>
        <button class="btn btn-ghost btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">View Proof</button>`;
    } else {
      const currentStatus = a.status === 'In Progress' ? 'In Progress' : 'Assigned';
      content = `<span style="font-size:11.5px;color:var(--muted);font-weight:600">${currentStatus}</span><span class="assign-not-submitted">Not submitted</span>`;
    }
    return `${content}<span class="assign-actions-trailing">${followBtn}${bellBtn}${moreBtn}</span>`;
  }

  function assignmentCard(a, received, history=false, principal=false) {
    const hasFollowup = !history && proofState(a) !== 'approved' && !isCancelled(a) && followupUnreadCount(a, received) > 0;
    const changesRequested=!history&&received&&proofState(a)==='declined'&&!isCancelled(a);
    const overdueCard=!history&&!received&&isAssignmentOverdue(a);
    const isNew=!history&&isNewAssignment(a);
    const expanded=expandedAssignmentId===a.id;
    const proofNotice='';
    const reminders=Math.max(0,Number(a.reminderCount||0));
    const reminderCount=!history&&received&&reminders?`<span class="assign-reminder-count">${reminders} reminder${reminders===1?'':'s'}</span>`:'';
    const changesBadge=changesRequested?'<span class="assign-changes-count">Changes requested</span>':'';
    const newReminder=!history&&hasNewReminder(a,received);
    const proofReady=!history&&!received&&awaitingApproval(a);
    const followupHistory=(followupThreadState(a)?.thread||[]).length;
    const submissionHistory=window._proofSubmissionHistory?.[String(a.appTaskId||'')]||[];
    const historyProofInline=history&&expanded&&submissionHistory.length
      ?`<div style="margin-top:9px">${renderNotificationProofs({...submissionHistory[0],taskId:String(a.appTaskId||'')})}</div>`:'';
    const historyLabels=history?`<div class="assign-history-labels">
      ${(a.proofSubmittedAt||proofState(a)==='approved')?'<span class="assign-history-label">Proof submitted</span>':''}
      ${followupHistory?`<span class="assign-history-label">${followupHistory} message${followupHistory===1?'':'s'}</span>`:'<span class="assign-history-label">No messages</span>'}
      <span class="assign-history-label">${reminders} reminder${reminders===1?'':'s'}</span>
    </div>`:'';
    return `<div class="wed-card assign-compact-card${hasFollowup ? ' has-followup' : ''}${newReminder?' has-new-reminder':''}${proofReady?' has-proof-ready':''}${changesRequested?' has-changes-requested':''}${overdueCard?' is-overdue':''}${expanded?' is-expanded':''}" data-assignment-id="${escapeHtml(String(a.id))}">
      <div class="wed-card-head assign-compact-head">
        <button type="button" class="assign-title-button" onclick="toggleAssignmentDetails('${a.id}')" aria-expanded="${expanded}">
          <span class="assign-expand-symbol">${expanded?'−':'+'}</span><span class="wed-card-title">${escapeHtml(a.title || '')}</span>${reminderCount}${changesBadge}
        </button>
        ${history?'':`<span class="dept-pill"><span class="dept-dot" style="background:${dcolor(a.dept)}"></span>${escapeHtml(a.dept || '')}</span>${pBadge(a.priority)}`}
      </div>
      <div class="wed-card-body assign-compact-body">
        ${historyLabels}
        ${history?`${(a.proofSubmittedAt||proofState(a)==='approved'||proofState(a)==='declined')?`<div class="assign-compact-summary"><div></div><div class="assign-actions"><button class="btn btn-ghost btn-sm" onclick="openProofReviewFromTasksTab('${a.id}')">View Proof</button></div></div>`:''}`:`<div class="assign-compact-summary">
          <div class="assign-card-meta">${dueDateBadge(a)}${isNew?'<span class="assign-task-new">New</span>':''}${proofNotice}</div>
          <div class="assign-actions">${assignmentActions(a, received, principal)}</div>
        </div>`}
        ${expanded?`<div class="assign-expanded-details">${assignmentDescription(a.summary, a.id)}${history?'':renderStepper(a)}${historyProofInline}</div>`:''}
      </div>
    </div>`;
  }

  window.toggleAssignmentDetails = function toggleAssignmentDetails(id) {
    expandedAssignmentId=expandedAssignmentId===id?null:id;
    if(expandedAssignmentId){
      const a=[...(tasksTabCache.assignedToMe||[]),...(tasksTabCache.assignedByMe||[]),...(tasksTabCache.overseenByMe||[])].find(row=>row.id===id);
      if(a){
        seenAssignmentStages.add(assignmentSeenKey(a));
        saveSeenStages(seenAssignmentStages);
        if(tasksTabMode==='received'){
          taskReminderSeenCounts[id]=Math.max(0,Number(a.reminderCount||0));
          a.recipientReminderSeenCount=taskReminderSeenCounts[id];
          saveTaskReminderSeenCounts();
          persistAssignmentSeen(a,{reminderCount:a.recipientReminderSeenCount});
        }
      }
    }
    renderTasksTabList();
  };

  // Create/update a shared assignment record in D1 via the Worker.
  // Fire-and-forget: failures are logged only, never block To Do/OneDrive writes.
  window.recordAssignment = async function recordAssignment(task) {
    if (!task || !task.email || !task.email.includes('@dhananipeg.com')) return;
    try {
      task.assignmentId = task.assignmentId || crypto.randomUUID();
      const userToken = await getAccessToken();
      const res=await fetch(`${fnBaseUrl()}/assignment`, {
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
          expectedVersion: task.assignmentVersion??null,
        }),
      });
      const data=await res.json().catch(()=>({}));
      if(res.status===409){toast(data.message||'This task changed. Reload before saving.');return false;}
      if(!res.ok)throw new Error(data.error||`HTTP ${res.status}`);
      task.assignmentVersion=Number(data.version||task.assignmentVersion||1);
      return true;
    } catch (err) {
      console.warn('Assignment record (D1) failed:', err.message);
      return false;
    }
  };

  window.setTasksTabMode = function setTasksTabMode(mode) {
    tasksTabMode = mode;
    if (mode === 'history') {
      (tasksTabCache.assignedToMe||[]).filter(a=>['Done','Cancelled'].includes(stageLabel(a))).forEach(a=>{
        seenAssignmentStages.add(assignmentSeenKey(a));
      });
      saveSeenStages(seenAssignmentStages);
    }
    document.getElementById('tasks-received-btn')?.classList.toggle('active', mode === 'received');
    document.getElementById('tasks-given-btn')?.classList.toggle('active', mode === 'given');
    document.getElementById('tasks-department-btn')?.classList.toggle('active', mode === 'department');
    document.getElementById('tasks-history-btn')?.classList.toggle('active', mode === 'history');
    const historyFilter=document.getElementById('tasks-history-filter');
    if(historyFilter)historyFilter.style.display=mode==='history'?'flex':'none';
    const desc = document.getElementById('tasks-tab-description');
    if (desc) {
      desc.textContent = mode === 'received'
        ? 'Your active assignments, status updates, proof, and conversations.'
        : mode === 'given'
        ? 'Tasks you assigned to others. Submitted proof needing review appears first.'
        : mode === 'department'
        ? 'Tasks in your department. You can monitor activity and participate in messages.'
        : tasksHistoryFilter==='cancelled'
        ? 'Cancelled tasks are kept here for reference.'
        : 'Approved and completed tasks, with their conversations and proof history.';
    }
    renderTasksTabList();
  };

  window.setTasksHistoryFilter = function setTasksHistoryFilter(filter) {
    tasksHistoryFilter=filter==='cancelled'?'cancelled':'completed';
    document.getElementById('tasks-history-completed-btn')?.classList.toggle('active',tasksHistoryFilter==='completed');
    document.getElementById('tasks-history-cancelled-btn')?.classList.toggle('active',tasksHistoryFilter==='cancelled');
    window.setTasksTabMode('history');
  };

  function assignmentsSignature(cache) {
    const sig = list => (list || [])
      .map(a => [a.id, a.status, a.proofStatus, a.summary, a.dueDate, a.title, a.dept, a.priority, a.updateAlertAt, a.reminderCount, a.assignerMessageSeenCount, a.recipientMessageSeenCount, a.recipientReminderSeenCount, a.version].join('|'))
      .join(';');
    return `${sig(cache?.assignedToMe)}::${sig(cache?.assignedByMe)}::${sig(cache?.overseenByMe)}`;
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
    const departmentBtn=document.getElementById('tasks-department-btn');
    if(departmentBtn)departmentBtn.style.display=(tasksTabCache.overseenByMe||[]).length?'':'none';
    window.updateNotificationCenter?.();
    renderTasksTabList();
  };

  window.getTasksNotificationSnapshot = function getTasksNotificationSnapshot() {
    return {
      assignedToMe: (tasksTabCache.assignedToMe || []).map(a => ({ ...a })),
      assignedByMe: (tasksTabCache.assignedByMe || []).map(a => ({ ...a })),
      overseenByMe: (tasksTabCache.overseenByMe || []).map(a => ({ ...a })),
    };
  };

  window.openAssignmentFromNotification = function openAssignmentFromNotification(assignmentId, received) {
    const mode = received ? 'received' : 'given';
    const list = received ? (tasksTabCache.assignedToMe || []) : (tasksTabCache.assignedByMe || []);
    const assignment = list.find(a => String(a.id) === String(assignmentId));
    if (!assignment) { toast('Task is no longer available'); return; }
    tasksTabMode = mode;
    expandedAssignmentId = assignment.id;
    const name = received ? assignment.assignerName : assignment.recipientName;
    const email = received ? assignment.assignerEmail : assignment.recipientEmail;
    tasksTabOpenGroups[mode].add(groupKey(name, email));
    nav('tasks');
    window.setTasksTabMode?.(mode);
    renderTasksTabList();
    setTimeout(() => {
      const card = document.querySelector(`[data-assignment-id="${CSS.escape(String(assignment.id))}"]`);
      card?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 80);
  };

  // Alerts the receiver of a new assignment (and the assignor of a new
  // submission, since that's also an unseen stage change) without requiring
  // the Tasks tab to be open: a badge on the sidebar "Tasks" nav item, and
  // one on each of the Received/Delegated toggle buttons.
  function updateTasksNavBadges() {
    const toMe = tasksTabCache.assignedToMe || [];
    const byMe = tasksTabCache.assignedByMe || [];
    const unseen = a => !seenAssignmentStages.has(assignmentSeenKey(a));
    const isPast = a => stageLabel(a) === 'Done' || stageLabel(a) === 'Cancelled';
    const notPast = a => !isPast(a);
    const receivedActive=toMe.filter(notPast);
    const givenActive=byMe.filter(notPast);
    const receivedGroups=groupAssignments(receivedActive,true);
    const givenGroups=groupAssignments(givenActive,false);
    const receivedEvents=receivedActive.filter(a=>isNewAssignment(a)||(proofState(a)==='declined'&&unseen(a))).length;
    const givenEvents=givenActive.filter(a=>awaitingApproval(a)&&unseen(a)).length;
    const receivedTransient=receivedGroups.reduce((sum,group)=>{
      const notices=groupTransientNotices('received',group);
      return sum+notices.followups+notices.reminders;
    },0);
    const givenTransient=givenGroups.reduce((sum,group)=>sum+groupTransientNotices('given',group).followups,0);
    const receivedCount=receivedEvents+receivedTransient;
    const givenCount=givenEvents+givenTransient;
    // Only the recipient is notified of approval/cancellation. The assignor
    // performed that action, so their copy must not create a History badge.
    const historyCount=toMe.filter(a=>isPast(a)&&unseen(a)).length;
    const setBadge = (id, count) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.textContent = count > 99 ? '99+' : count;
      el.style.display = count > 0 ? '' : 'none';
    };
    setBadge('nb-tasks', receivedCount + givenCount + historyCount);
    setBadge('tasks-received-badge', receivedCount);
    setBadge('tasks-given-badge', givenCount);
    setBadge('tasks-history-badge', historyCount);
  }

  // Exposed so index.html's notification poll can force a repaint when only
  // the KV follow-up thread changed — renderMyTasks(true) alone won't redraw,
  // since its D1-row signature check has nothing to compare there (see
  // checkAndLoadProofNotifications's followup-signature diff).
  window.renderTasksTabList = renderTasksTabList;
  function renderTasksTabList() {
    const tb = document.getElementById('tasks-tbody');
    if (!tb) return;
    updateTasksNavBadges();
    const mode = tasksTabMode;
    const { list, received, principal, sortFn } = tasksTabModeSource(mode);
    if (!list.length) {
      const emptyText = mode === 'history'
        ? (tasksHistoryFilter==='cancelled'?'No cancelled tasks':'No completed tasks yet')
        : 'No tasks yet';
      tb.innerHTML = `<div class="empty-state"><div class="es-text">${emptyText}</div></div>`;
      return;
    }
    const openGroups = tasksTabOpenGroups[mode];
    const groups=groupAssignments(list, received, sortFn);
    tb.innerHTML = groups.map(group => {
      const open = openGroups.has(group.key);
      const noun = group.items.length === 1 ? 'task' : 'tasks';
      const summaryText = mode==='history'?stageSummary(group.items):friendlyGroupSummary(group.items,group.received);
      const safeGroupKey = escapeHtml(JSON.stringify(group.key));
      const newCount=group.items.filter(isNewAssignment).length;
      const newCountBadge=newCount?`<span class="assign-person-new-count">${newCount>99?'99+':newCount}</span>`:'';
      const changesRequestedCount=mode==='received'?group.items.filter(a=>proofState(a)==='declined'&&!isCancelled(a)).length:0;
      const changesGroupBadge=changesRequestedCount?`<span class="assign-changes-count">${changesRequestedCount} change${changesRequestedCount===1?'':'s'} requested</span>`:'';
      const proofCount=mode==='given'?group.items.filter(awaitingApproval).length:0;
      const proofGroupBadge=proofCount?`<span class="assign-approval-badge">${proofCount} proof${proofCount===1?'':'s'}</span>`:'';
      const transientNotices=groupTransientNotices(mode,group);
      const unreadFollowups=transientNotices.followups;
      const newReminders=transientNotices.reminders;
      const followupGroupBadge=unreadFollowups?'<span class="assign-followup-group-badge">1 unread</span>':'';
      const personNotices=unreadFollowups+newReminders;
      const avatarNotification=personNotices?`<span class="assign-new-badge">${personNotices>99?'99+':personNotices}</span>`:'';
      const reminderGroupBadge=newReminders?`<span class="assign-alert-group-badge">${newReminders} reminder${newReminders===1?'':'s'}</span>`:'';
      const toItems=mode==='history'?group.items.filter(a=>!a._received):[];
      const byItems=mode==='history'?group.items.filter(a=>a._received):[];
      const defaultDirection=toItems.length?'to':'by';
      const historyDirection=mode==='history'?(tasksHistoryDirections.get(group.key)||defaultDirection):'';
      const shownItems=mode==='history'?(historyDirection==='by'?byItems:toItems):group.items;
      const pageKey=mode==='history'?`${group.key}::${historyDirection}`:group.key;
      const visibleCount=tasksTabVisibleCounts[mode].get(pageKey)||TASKS_PAGE_SIZE;
      const visibleItems=shownItems.slice(0,visibleCount);
      const remaining=Math.max(0,shownItems.length-visibleItems.length);
      const directionToggle=mode==='history'?`<div class="assign-history-direction">
        <button type="button" class="${historyDirection==='to'?'active':''}" onclick="setTasksHistoryDirection(event,${safeGroupKey},'to')" ${toItems.length?'':'disabled'}>Assigned to ${escapeHtml(group.name)} <span>${toItems.length}</span></button>
        <button type="button" class="${historyDirection==='by'?'active':''}" onclick="setTasksHistoryDirection(event,${safeGroupKey},'by')" ${byItems.length?'':'disabled'}>Assigned by ${escapeHtml(group.name)} <span>${byItems.length}</span></button>
      </div>`:'';
      const cards = open
        ? `${directionToggle}<div class="assign-cards">${visibleItems.map(a => assignmentCard(a, mode==='history'?a._received:group.received,mode==='history',principal)).join('')}${remaining?`<button type="button" class="assign-show-more" onclick="showMoreAssignments(${JSON.stringify(pageKey)})">Show 10 more <span>(${remaining} remaining)</span></button>`:''}</div>`
        : '';
      const groupName = escapeHtml(group.name);
      const historySummary=mode==='history'
        ?`${group.items.length} task${group.items.length===1?'':'s'} · ${toItems.length} assigned to · ${byItems.length} assigned by`
        :summaryText;
      return `<div class="assign-group">
        <div class="assign-group-head" onclick="toggleTasksGroup(${safeGroupKey})">
          <span class="assign-group-toggle">${open ? '−' : '+'}</span>
          <span class="assign-avatar-wrap">${av(group.name, 24)}${avatarNotification}</span>
          <span class="assign-group-name">${groupName}</span>
          ${newCountBadge}
          ${changesGroupBadge}
          ${proofGroupBadge}
          ${followupGroupBadge}
          ${reminderGroupBadge}
          <span class="assign-group-summary">${escapeHtml(historySummary||`${group.items.length} ${noun}`)}</span>
        </div>
        ${cards}
      </div>`;
    }).join('');
    syncAssignDescClamped(tb);
  }

  window.showMoreAssignments = function showMoreAssignments(key) {
    const map=tasksTabVisibleCounts[tasksTabMode];
    map.set(key,(map.get(key)||TASKS_PAGE_SIZE)+TASKS_PAGE_SIZE);
    renderTasksTabList();
  };

  window.setTasksHistoryDirection = function setTasksHistoryDirection(event,key,direction) {
    event?.stopPropagation();
    tasksHistoryDirections.set(key,direction==='by'?'by':'to');
    renderTasksTabList();
  };

  window.toggleTasksGroup = function toggleTasksGroup(key) {
    const openGroups = tasksTabOpenGroups[tasksTabMode];
    if (openGroups.has(key)) {
      openGroups.delete(key);
    } else {
      openGroups.add(key);
      const { list, received, sortFn } = tasksTabModeSource(tasksTabMode);
      const group = groupAssignments(list, received, sortFn).find(item => item.key === key);
      if (group && tasksTabMode !== 'history') {
        groupNoticeTotals[groupNoticeKey(tasksTabMode,key)] = {
          followups: incomingFollowupTotal(group.items,group.received),
          reminders: tasksTabMode==='received'
            ? group.items.reduce((sum,a)=>sum+Math.max(0,Number(a.reminderCount||0)),0)
            : 0,
        };
        saveGroupNoticeTotals();
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

  // One-click "update required" nudge — Delegated cards only. Independent of
  // the follow-up thread: just flips a timestamp flag on the assignment row.
  window.sendUpdateAlert = async function sendUpdateAlert(id) {
    const a = (tasksTabCache.assignedByMe || []).find(x => x.id === id);
    if (!a) return;
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignment-alert`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({ id }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const { updateAlertAt, reminderCount, version, updatedAt } = await res.json();
      a.updateAlertAt = updateAlertAt || new Date().toISOString();
      a.reminderCount = Number(reminderCount || 0);
      if (version) a.version = Number(version);
      if (updatedAt) a.updatedAt = updatedAt;
      renderTasksTabList();
      toast(`Reminder ${a.reminderCount} sent to ${a.recipientName||a.recipientEmail||'the assignee'}`);
    } catch (err) {
      console.warn('Send update alert failed:', err.message);
      toast('Could not send alert — try again');
    }
  };

  // Manual dismiss for the recipient (or the assigner, retracting their own
  // alert) — see the comment on alertLabel for why this can't rely solely
  // on the automatic clear-on-status-change path.
  window.dismissUpdateAlert = async function dismissUpdateAlert(id) {
    const a = (tasksTabCache.assignedToMe || []).find(x => x.id === id)
      || (tasksTabCache.assignedByMe || []).find(x => x.id === id);
    if (!a) return;
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignment-alert-clear`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({ id }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      a.updateAlertAt = null;
      renderTasksTabList();
      toast('Alert dismissed');
    } catch (err) {
      console.warn('Dismiss alert failed:', err.message);
      toast('Could not dismiss alert — try again');
    }
  };

  // Shared single dropdown (same pattern as the Outlook right-click menu)
  // repositioned per click rather than one dropdown DOM per card — there can
  // be dozens of Delegated cards on screen at once.
  let _moreMenuAssignmentId = null;
  window.openAssignMoreMenu = function openAssignMoreMenu(e, id) {
    e.preventDefault(); e.stopPropagation();
    _moreMenuAssignmentId = id;
    const m = document.getElementById('assign-more-menu');
    if (!m) return;
    const rect = e.currentTarget.getBoundingClientRect();
    const x = Math.min(rect.left, window.innerWidth - 180);
    const y = Math.min(rect.bottom + 4, window.innerHeight - 60);
    m.style.left = x + 'px'; m.style.top = y + 'px'; m.style.display = 'block';
  };
  function hideAssignMoreMenu() {
    const m = document.getElementById('assign-more-menu');
    if (m) m.style.display = 'none';
  }
  document.addEventListener('click', hideAssignMoreMenu);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') hideAssignMoreMenu(); });

  // Assigner-only. Confirms, then optionally collects a reason, before
  // calling off the task — mirrors the existing "Request Changes" reason
  // pattern. The recipient's copy is notified both in-app (status flips to
  // Cancelled next refresh) and by email, same as a declined proof today.
  window.cancelAssignmentPrompt = async function cancelAssignmentPrompt() {
    hideAssignMoreMenu();
    const id = _moreMenuAssignmentId;
    const a = (tasksTabCache.assignedByMe || []).find(x => x.id === id);
    if (!a) return;
    const who = a.recipientName || a.recipientEmail || 'the assignee';
    if (!confirm(`Cancel "${a.title}"?\n\nIt will be removed from active tasks and moved to Cancelled History for both users.`)) return;
    const reason = '';
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignment-cancel`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({ id, reason }),
      });
      if (!res.ok) {
        const detail = await res.json().catch(() => ({}));
        throw new Error(detail.error || `HTTP ${res.status}`);
      }
      const { cancelledAt } = await res.json();
      a.status = 'Cancelled';
      a.cancelReason = reason;
      a.cancelledAt = cancelledAt || new Date().toISOString();
      a.updatedAt = a.cancelledAt;
      a.updateAlertAt = null;
      const localTask=(typeof tasks!=='undefined'?tasks:[]).find(t=>String(t.id)===String(a.appTaskId));
      if(localTask){localTask.status='Cancelled';localTask.cancelledAt=a.cancelledAt;localTask.cancelReason=reason;await saveTasksToOneDrive();refreshAll();}
      renderTasksTabList();
      toast(`Task cancelled — notifying ${who}`);
      window.sendTaskCancelledEmail?.(a);
    } catch (err) {
      console.warn('Cancel assignment failed:', err.message);
      toast('Could not cancel task: ' + err.message);
    }
  };

  window.cancelAssignmentDirect = function cancelAssignmentDirect(id) {
    _moreMenuAssignmentId=id;
    window.cancelAssignmentPrompt();
  };

  // Opens the generic task follow-up modal (index.html: showTaskFollowupModal),
  // available on both Received and Delegated cards regardless of proof status.
  window.openTaskFollowup = function openTaskFollowup(id, received) {
    const principal=received==='principal';
    const list = principal ? (tasksTabCache.overseenByMe || []) : received ? (tasksTabCache.assignedToMe || []) : (tasksTabCache.assignedByMe || []);
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
      role: principal ? 'principal' : received ? 'assignee' : 'assignor',
    });
  };

  // Called by showTaskFollowupModal once the thread has loaded, so opening
  // the modal clears the unread count the same way expanding a group clears
  // its "new" badge. Stores the thread length seen so far (not just a flag)
  // so a later re-open can compute exactly how many new messages arrived.
  // role must match the 'assignee'/'assignor' key format used by
  // followupSeenKey above — see the comment there for why role is part of
  // the key (self-assigned tasks share one assignment id across both views).
  window.markTaskFollowupSeen = function markTaskFollowupSeen(assignmentId, threadLen, role) {
    if (!assignmentId) return;
    const key = `${assignmentId}::${role === 'assignor' ? 'assignor' : 'assignee'}`;
    followupSeenLengths[key] = threadLen;
    saveFollowupSeenLengths(followupSeenLengths);
    const a=[...(tasksTabCache.assignedToMe||[]),...(tasksTabCache.assignedByMe||[])].find(row=>row.id===assignmentId);
    if(a){
      if(role==='assignor')a.assignerMessageSeenCount=threadLen;
      else a.recipientMessageSeenCount=threadLen;
      persistAssignmentSeen(a,{threadLen,reminderCount:role==='assignee'?Number(a.recipientReminderSeenCount||0):0});
    }
    renderTasksTabList();
  };

  window.openProofReviewFromTasksTab = async function openProofReviewFromTasksTab(id) {
    const a = [...(tasksTabCache.assignedByMe || []),...(tasksTabCache.assignedToMe || [])].find(x => x.id === id);
    if (!a) return;
    if (typeof window.openTaskProofReview !== 'function') {
      toast('Proof review is still loading — try again');
      return;
    }
    seenAssignmentStages.add(assignmentSeenKey(a));
    saveSeenStages(seenAssignmentStages);
    renderTasksTabList();
    await window.openTaskProofReview(a);
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
    const assignment = (tasksTabCache.assignedToMe || []).find(a => a.id === id);
    if (!assignment) return;
    try {
      const userToken = await getAccessToken();
      const res = await fetch(`${fnBaseUrl()}/assignment-status`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${userToken}` },
        body: JSON.stringify({ id, status, expectedVersion: assignment.version }),
      });
      if (res.status === 409) {
        await renderMyTasks(false);
        toast('This task changed elsewhere. Latest version loaded — please try again.');
        return;
      }
      if(res.status===423){
        const data=await res.json().catch(()=>({}));
        await renderMyTasks(true);
        toast(`${data.editorName||'Another user'} is currently editing this task. Please wait.`);
        return;
      }
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const { version, updatedAt } = await res.json();
      const changedRows = [
        ...(tasksTabCache.assignedToMe || []).filter(a => a.id === id),
        ...(tasksTabCache.assignedByMe || []).filter(a => a.id === id),
      ];
      changedRows.forEach(row => {
        row.status = status;
        row.version = version || row.version;
        row.updatedAt = updatedAt || row.updatedAt;
        // The assignee just acted on this task, so the original assignment
        // is no longer new. In Progress intentionally shares Assigned's key.
        seenAssignmentStages.add(assignmentSeenKey(row));
      });
      saveSeenStages(seenAssignmentStages);
      renderTasksTabList();
      toast('Progress updated');
    } catch (err) {
      console.warn('Update assignment status failed:', err.message);
      toast('Could not update progress — try again');
      renderTasksTabList(); // revert the <select> to the last-known-good cached value
    }
  };
})();
