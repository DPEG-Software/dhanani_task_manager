(function () {
  function contactEmailValue(item) {
    if (!item) return "";
    if (typeof item === "string") return item;
    return item.address || item.emailAddress?.address || item.email || item.mail || item.userPrincipalName || "";
  }

  function contactNameValue(item, email) {
    if (item?.displayName) return item.displayName;
    if (item?.name) return item.name;
    if (item?.emailAddress?.name) return item.emailAddress.name;
    if (item?.givenName && item?.surname) return `${item.givenName} ${item.surname}`;
    return String(email || "").split("@")[0] || "Unknown";
  }

  function upsertContact(raw, defaultDept) {
    const candidates = [
      ...(raw.emailAddresses || []),
      ...(raw.scoredEmailAddresses || []),
      raw.mail,
      raw.userPrincipalName,
      raw.email,
      raw.emailAddress,
    ];
    let email = "";
    for (const c of candidates) {
      email = normEmail(contactEmailValue(c));
      if (email && email.includes("@")) break;
    }
    if (!email || !email.includes("@")) return false;

    const name = (raw.displayName || raw.name || raw.emailAddress?.name || contactNameValue(raw, email)).trim();
    const dept = raw.department || defaultDept || (isInternalEmail(email) ? "Needs Department" : "Outside DPEG");
    const role = raw.jobTitle || raw.role || "";
    const key = staffKey(email, name);
    const before = JSON.stringify(staffConfig[key] || null);
    staffConfig[key] = {
      ...(staffConfig[key] || {}),
      name,
      email,
      dept: staffConfig[key]?.dept || dept,
      role: staffConfig[key]?.role || role,
    };
    return JSON.stringify(staffConfig[key]) !== before;
  }

  async function graphGetAll(url, token, maxPages) {
    const rows = [];
    let next = url;
    let pages = 0;
    while (next && pages < (maxPages || 25)) {
      const res = await fetch(next, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) break;
      const data = await res.json();
      rows.push(...(data.value || []));
      next = data["@odata.nextLink"] || "";
      pages++;
    }
    return rows;
  }

  async function syncOutlookContactsFull(token) {
    let changed = 0;
    const select = "$select=displayName,emailAddresses,department,jobTitle";
    const contacts = await graphGetAll(`https://graph.microsoft.com/v1.0/me/contacts?$top=999&${select}`, token, 50);
    contacts.forEach((c) => { if (upsertContact(c)) changed++; });

    const folders = await graphGetAll("https://graph.microsoft.com/v1.0/me/contactFolders?$top=100&$select=id,displayName", token, 10);
    for (const folder of folders) {
      const folderContacts = await graphGetAll(`https://graph.microsoft.com/v1.0/me/contactFolders/${folder.id}/contacts?$top=999&${select}`, token, 25);
      folderContacts.forEach((c) => { if (upsertContact(c)) changed++; });
    }
    return changed;
  }

  async function syncPeopleSuggestions(token) {
    let changed = 0;
    const people = await graphGetAll(
      "https://graph.microsoft.com/v1.0/me/people?$top=999&$select=displayName,emailAddresses,scoredEmailAddresses,department,jobTitle,userPrincipalName",
      token,
      25
    );
    people.forEach((p) => { if (upsertContact(p)) changed++; });
    return changed;
  }

  function contactsFromMessage(m) {
    return [
      m.from?.emailAddress,
      m.sender?.emailAddress,
      ...(m.toRecipients || []).map((r) => r.emailAddress),
      ...(m.ccRecipients || []).map((r) => r.emailAddress),
      ...(m.bccRecipients || []).map((r) => r.emailAddress),
    ].filter(Boolean);
  }

  async function syncMailboxNames(token) {
    let changed = 0;
    const urls = [
      "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=250&$select=from,sender,toRecipients,ccRecipients,receivedDateTime&$orderby=receivedDateTime desc",
      "https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$top=250&$select=from,sender,toRecipients,ccRecipients,bccRecipients,sentDateTime&$orderby=sentDateTime desc",
    ];
    for (const url of urls) {
      const messages = await graphGetAll(url, token, 5);
      messages.forEach((m) => {
        contactsFromMessage(m).forEach((c) => {
          if (upsertContact({ displayName: c.name, emailAddress: c }, isInternalEmail(c.address) ? "Needs Department" : "Outside DPEG")) changed++;
        });
      });
    }
    return changed;
  }

  function normalizedSearchToken(token) {
    return String(token || "")
      .trim()
      .toLowerCase()
      .replace(/^[<"'(\[\s]+|[>:"')\]\s]+$/g, "");
  }

  function profileContacts() {
    const map = new Map();
    if (typeof PRINCIPALS !== "undefined") {
      Object.entries(PRINCIPALS).forEach(([email, profile]) => {
        map.set(normEmail(email), {
          name: profile.name || email.split("@")[0],
          email: normEmail(email),
          dept: "Needs Department",
          role: profile.role || "",
        });
      });
    }
    if (typeof ADMIN_EMAILS !== "undefined" && Array.isArray(ADMIN_EMAILS)) {
      ADMIN_EMAILS.forEach((email) => {
        const key = normEmail(email);
        if (!key || map.has(key)) return;
        const configured = Object.values(staffConfig).find((p) => normEmail(p.email) === key);
        map.set(key, {
          name: configured?.name || key.split("@")[0].replace(/[._-]+/g, " "),
          email: key,
          dept: configured?.dept || "Needs Department",
          role: configured?.role || "Admin",
        });
      });
    }
    if (currentUser?.email) {
      const key = normEmail(currentUser.email);
      const existing = map.get(key) || {};
      map.set(key, {
        ...existing,
        name: currentUser.name || existing.name || key.split("@")[0],
        email: key,
        dept: existing.dept || "Needs Department",
        role: existing.role || "Current user",
      });
    }
    return [...map.values()];
  }

  function contactSearchPool() {
    const map = new Map();
    profileContacts().forEach((p) => {
      if (p?.email) map.set(normEmail(p.email), p);
    });
    Object.values(staffConfig).forEach((p) => {
      if (!p?.email || !p?.name) return;
      const key = normEmail(p.email);
      map.set(key, { ...(map.get(key) || {}), ...p, email: key });
    });
    return [...map.values()];
  }

  function contactMatches(token) {
    const q = normalizedSearchToken(token);
    if (q.length < 2) return [];
    const seen = new Set();
    return contactSearchPool()
      .filter((p) => p?.email && p?.name)
      .filter((p) => {
        const email = normEmail(p.email);
        if (!email || seen.has(email)) return false;
        seen.add(email);
        const name = String(p.name || "").toLowerCase();
        const role = String(p.role || "").toLowerCase();
        const dept = String(p.dept || "").toLowerCase();
        return name.startsWith(q) || email.startsWith(q) || name.includes(q) || email.includes(q) || role.includes(q) || dept.includes(q);
      })
      .sort((a, b) => {
        const an = String(a.name || "").toLowerCase();
        const bn = String(b.name || "").toLowerCase();
        const ae = normEmail(a.email);
        const be = normEmail(b.email);
        const ar = (an.startsWith(q) || ae.startsWith(q)) ? 0 : 1;
        const br = (bn.startsWith(q) || be.startsWith(q)) ? 0 : 1;
        return ar - br || an.localeCompare(bn);
      })
      .slice(0, 20);
  }

  function renderContactItems(acId, matches, selectCall) {
    return matches.map((p) => `
      <div class="compose-ac-item"
        onmousedown="event.preventDefault();${selectCall(p)}"
        onmouseover="document.querySelectorAll('#${acId} .compose-ac-item').forEach(x=>x.classList.remove('ac-focused'));this.classList.add('ac-focused')">
        ${av(p.name || "?", 30)}
        <div style="flex:1;min-width:0;overflow:hidden">
          <div class="compose-ac-name">${p.name || p.email}</div>
          <div class="compose-ac-email">${p.email || ""}</div>
          ${p.role || p.dept ? `<div class="compose-ac-role">${p.role || p.dept}</div>` : ""}
        </div>
      </div>`).join("");
  }

  window.syncContacts = async function syncContacts() {
    if (!isAdmin()) { toast("Admin access only"); return; }
    const btn = document.getElementById("sync-contacts-btn");
    const status = document.getElementById("sync-status");
    if (btn) { btn.disabled = true; btn.textContent = "Syncing..."; }
    if (status) status.textContent = "Pulling Outlook contacts...";
    try {
      let token;
      try {
        token = (await msalInstance.acquireTokenSilent({ scopes: SCOPES_CONTACTS, account: currentAccount })).accessToken;
      } catch {
        token = (await msalInstance.acquireTokenPopup({ scopes: SCOPES_CONTACTS })).accessToken;
      }
      let changed = 0;
      changed += await syncOutlookContactsFull(token);
      changed += await syncPeopleSuggestions(token);
      changed += await syncMailboxNames(token);
      await saveTasksToOneDrive();
      renderAdminPeopleList();
      renderAdminDeptEditor();
      initSelects();
      const available = allKnownPeople().length;
      if (status) {
        status.textContent = `Synced ${changed} new/updated contact${changed !== 1 ? "s" : ""}. ${available} available in autocomplete.`;
        status.style.color = "var(--forest)";
      }
      toast(`${available} Outlook contacts available`);
    } catch (err) {
      if (status) {
        status.textContent = `Sync failed: ${err.message}`;
        status.style.color = "#dc2626";
      }
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = "Sync from Outlook"; }
    }
  };

  window.autoSyncContacts = async function autoSyncContacts() {
    if (!isAdmin()) return;
    try {
      const token = (await msalInstance.acquireTokenSilent({ scopes: SCOPES_CONTACTS, account: currentAccount })).accessToken;
      const before = JSON.stringify(staffConfig);
      await syncOutlookContactsFull(token);
      await syncPeopleSuggestions(token);
      if (JSON.stringify(staffConfig) !== before) await saveTasksToOneDrive();
    } catch {}
  };

  window.showComposeAC = function showComposeAC(inputId, acId) {
    const input = document.getElementById(inputId);
    const ac = document.getElementById(acId);
    if (!input || !ac) return;
    const val = input.value;
    const lastSep = Math.max(val.lastIndexOf(","), val.lastIndexOf(";"));
    const token = normalizedSearchToken(lastSep >= 0 ? val.slice(lastSep + 1) : val);
    const matches = contactMatches(token);
    if (!matches.length) { ac.style.display = "none"; return; }
    ac.innerHTML = renderContactItems(acId, matches, (p) => `selectComposeAC('${inputId}','${acId}','${String(p.email || "").replace(/'/g, "\\'")}')`);
    const rect = input.getBoundingClientRect();
    ac.style.position = "fixed";
    ac.style.left = `${rect.left}px`;
    ac.style.top = `${rect.bottom + 2}px`;
    ac.style.width = `${rect.width}px`;
    ac.style.right = "auto";
    ac.style.display = "block";
  };

  window.showAddTaskAC = function showAddTaskAC(val) {
    const ac = document.getElementById("nt-ac");
    const input = document.getElementById("nt-person");
    const clearBtn = document.getElementById("nt-clear-btn");
    if (clearBtn) clearBtn.style.display = val ? "block" : "none";
    if (!ac || !input) return;
    const matches = contactMatches(normalizedSearchToken(val));
    if (!matches.length) { ac.style.display = "none"; return; }
    ac.innerHTML = renderContactItems("nt-ac", matches, (p) =>
      `selectAddTaskAC('${String(p.name || "").replace(/'/g, "\\'")}','${String(p.email || "").replace(/'/g, "\\'")}','${String(p.dept || "").replace(/'/g, "\\'")}')`
    );
    const rect = input.getBoundingClientRect();
    ac.style.position = "fixed";
    ac.style.left = `${rect.left}px`;
    ac.style.top = `${rect.bottom + 2}px`;
    ac.style.width = `${rect.width}px`;
    ac.style.right = "auto";
    ac.style.display = "block";
  };
})();
