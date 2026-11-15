(() => {
  "use strict";

  const CLIENT_ID_KEY = "mailnotes.microsoftClientId";
  const DEFAULT_CLIENT_ID = "3143ac7e-aff2-4cc3-9bcd-54126b9a7d4d";
  const GRAPH_ORIGIN = "https://graph.microsoft.com";
  const GRAPH_ROOT = `${GRAPH_ORIGIN}/v1.0`;
  const MAIL_SCOPE = "Mail.Read";
  const FOLDER_PAGE_SIZE = 100;
  const MESSAGE_PAGE_SIZE = 25;

  const elements = {
    accountLabel: document.getElementById("account-label"),
    clientId: document.getElementById("client-id"),
    connectButton: document.getElementById("connect-button"),
    connectPanel: document.getElementById("connect-panel"),
    copyButton: document.getElementById("copy-button"),
    folderSelect: document.getElementById("folder-select"),
    listStatus: document.getElementById("list-status"),
    loadMoreButton: document.getElementById("load-more-button"),
    mailPanel: document.getElementById("mail-panel"),
    messageList: document.getElementById("message-list"),
    refreshButton: document.getElementById("refresh-button"),
    searchInput: document.getElementById("search-input"),
    selectAll: document.getElementById("select-all"),
    selectionCount: document.getElementById("selection-count"),
    signOutButton: document.getElementById("sign-out"),
    statusMessage: document.getElementById("status-message"),
  };

  const state = {
    client: null,
    account: null,
    token: null,
    folders: [],
    messages: [],
    selectedIds: new Set(),
    nextLink: null,
    loading: false,
    searchTimer: null,
  };

  function showStatus(message, kind = "info") {
    elements.statusMessage.textContent = message;
    elements.statusMessage.dataset.kind = kind;
    elements.statusMessage.hidden = false;
  }

  function clearStatus() {
    elements.statusMessage.textContent = "";
    elements.statusMessage.hidden = true;
    delete elements.statusMessage.dataset.kind;
  }

  function setBusy(busy, button) {
    state.loading = busy;
    elements.connectButton.disabled = busy;
    elements.refreshButton.disabled = busy;
    elements.loadMoreButton.disabled = busy;
    elements.folderSelect.disabled = busy;
    if (button) button.disabled = busy;
  }

  function makeClient(clientId) {
    if (!window.msal || !window.msal.PublicClientApplication) {
      throw new Error("Microsoft sign-in could not load. Check your internet connection and reload this page.");
    }

    return new window.msal.PublicClientApplication({
      auth: {
        clientId,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: `${window.location.origin}${window.location.pathname}`,
      },
      cache: {
        cacheLocation: "sessionStorage",
      },
    });
  }

  async function getAccessToken(interactive = false) {
    if (!state.client || !state.account) {
      throw new Error("Connect your Outlook account to continue.");
    }

    try {
      const result = await state.client.acquireTokenSilent({
        account: state.account,
        scopes: [MAIL_SCOPE],
      });
      state.token = result.accessToken;
      return state.token;
    } catch (error) {
      if (!interactive || !(error instanceof window.msal.InteractionRequiredAuthError)) {
        throw error;
      }

      const result = await state.client.acquireTokenPopup({
        account: state.account,
        scopes: [MAIL_SCOPE],
      });
      state.token = result.accessToken;
      return state.token;
    }
  }

  async function graphGet(url) {
    const token = await getAccessToken();
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.body-content-type="html"',
      },
    });

    if (response.status === 401) {
      state.token = null;
      throw new Error("Your Microsoft sign-in expired. Sign out and connect again.");
    }
    if (!response.ok) {
      let details = "";
      try {
        const payload = await response.json();
        details = payload.error && payload.error.message ? ` ${payload.error.message}` : "";
      } catch {
        details = "";
      }
      throw new Error(`Microsoft Graph returned an error (${response.status}).${details}`);
    }

    return response.json();
  }

  function validateNextLink(link) {
    const url = new URL(link);
    if (url.origin !== GRAPH_ORIGIN) {
      throw new Error("Microsoft returned an unexpected paging link. Refresh the message list to continue.");
    }
    if (!url.pathname.startsWith("/v1.0/me/")) {
      throw new Error("Microsoft returned an unexpected paging link. Refresh the message list to continue.");
    }
    return url.href;
  }

  function formatDate(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return new Intl.DateTimeFormat(undefined, {
      month: "short",
      day: "numeric",
      year: date.getFullYear() === new Date().getFullYear() ? undefined : "numeric",
    }).format(date);
  }

  function cleanText(value) {
    return String(value || "")
      .replace(/\r\n?/g, "\n")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF\u00AD\u034F]/g, "")
      .replace(/\u00A0/g, " ")
      .replace(/\[[^\]\n]{30,}\]\s*<https?:\/\/[^>\s]+>/gi, "")
      .replace(/<https?:\/\/[^>\s]+>/gi, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/[ \t]{2,}/g, " ")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function htmlToText(html) {
    const parsed = new DOMParser().parseFromString(html, "text/html");
    parsed.querySelectorAll("script, style, noscript, img, svg, picture, head").forEach((node) => node.remove());

    function extract(node) {
      if (node.nodeType === Node.TEXT_NODE) return node.nodeValue || "";
      if (node.nodeType !== Node.ELEMENT_NODE) return "";

      const element = node;
      const tag = element.tagName.toLowerCase();
      if (tag === "br" || tag === "hr") return "\n";

      const content = Array.from(element.childNodes, extract).join("");
      if (tag === "td" || tag === "th") return `${content} `;
      if (/^(address|article|blockquote|div|footer|h[1-6]|li|main|ol|p|section|table|tr|ul)$/.test(tag)) {
        return `\n${content}\n`;
      }
      return content;
    }

    return cleanText(extract(parsed.body));
  }

  function normalizedBody(message) {
    const body = message.body && message.body.content ? message.body.content : "";
    if (message.body && message.body.contentType === "html") {
      return htmlToText(body);
    }
    return cleanText(body);
  }

  function searchableText(message) {
    const sender = message.from && message.from.emailAddress
      ? `${message.from.emailAddress.name || ""} ${message.from.emailAddress.address || ""}`
      : "";
    return `${message.subject || ""} ${sender} ${message.bodyPreview || ""}`.toLowerCase();
  }

  function filteredMessages() {
    const query = elements.searchInput.value.trim().toLowerCase();
    if (!query) return state.messages;
    return state.messages.filter((message) => searchableText(message).includes(query));
  }

  function createMessageRow(message) {
    const row = document.createElement("label");
    row.className = "message-row";
    row.classList.toggle("is-selected", state.selectedIds.has(message.id));

    const checkbox = document.createElement("input");
    checkbox.className = "message-check";
    checkbox.type = "checkbox";
    checkbox.checked = state.selectedIds.has(message.id);
    checkbox.setAttribute("aria-label", `Select ${message.subject || "email"}`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.selectedIds.add(message.id);
      else state.selectedIds.delete(message.id);
      row.classList.toggle("is-selected", checkbox.checked);
      updateSelection();
    });

    const sender = message.from && message.from.emailAddress ? message.from.emailAddress : {};
    const senderName = sender.name || sender.address || "Unknown sender";
    const avatar = document.createElement("span");
    avatar.className = "sender-avatar";
    avatar.setAttribute("aria-hidden", "true");
    avatar.textContent = senderName.trim().charAt(0).toUpperCase() || "?";

    const main = document.createElement("span");
    main.className = "message-main";

    const top = document.createElement("span");
    top.className = "message-top";
    const from = document.createElement("span");
    from.className = "message-from";
    from.textContent = senderName;
    const date = document.createElement("span");
    date.className = "message-date";
    date.textContent = formatDate(message.receivedDateTime);
    top.append(from, date);

    const subject = document.createElement("span");
    subject.className = "message-subject";
    subject.textContent = message.subject || "(No subject)";

    const preview = document.createElement("span");
    preview.className = "message-preview";
    preview.textContent = cleanText(message.bodyPreview || "").replace(/\s+/g, " ");

    main.append(top, subject, preview);
    row.append(checkbox, avatar, main);
    return row;
  }

  function renderMessages() {
    const visibleMessages = filteredMessages();
    elements.messageList.replaceChildren();

    if (visibleMessages.length === 0) {
      const empty = document.createElement("p");
      empty.className = "empty-state";
      empty.textContent = state.messages.length === 0
        ? "No messages found in this folder."
        : "No messages match your search.";
      elements.messageList.append(empty);
    } else {
      const fragment = document.createDocumentFragment();
      visibleMessages.forEach((message) => fragment.append(createMessageRow(message)));
      elements.messageList.append(fragment);
    }

    const visibleIds = new Set(visibleMessages.map((message) => message.id));
    const visibleSelected = visibleMessages.filter((message) => state.selectedIds.has(message.id)).length;
    elements.selectAll.checked = visibleMessages.length > 0 && visibleSelected === visibleMessages.length;
    elements.selectAll.indeterminate = visibleSelected > 0 && visibleSelected < visibleMessages.length;
    elements.selectionCount.textContent = `${state.selectedIds.size} selected`;
    elements.copyButton.disabled = state.selectedIds.size === 0 || state.loading;
    elements.listStatus.textContent = state.messages.length
      ? `${state.messages.length} loaded${visibleIds.size !== state.messages.length ? ` · ${visibleIds.size} shown` : ""}`
      : "";
  }

  function updateSelection() {
    renderMessages();
  }

  function setSignedIn(account) {
    state.account = account;
    elements.connectPanel.hidden = true;
    elements.mailPanel.hidden = false;
    elements.signOutButton.hidden = false;
    elements.accountLabel.hidden = !account.username;
    elements.accountLabel.textContent = account.username || "";
  }

  async function loadFolders() {
    const payload = await graphGet(`${GRAPH_ROOT}/me/mailFolders?$top=${FOLDER_PAGE_SIZE}&$select=id,displayName,totalItemCount`);
    const folders = Array.isArray(payload.value) ? payload.value : [];
    state.folders = folders;

    const fragment = document.createDocumentFragment();
    folders
      .sort((left, right) => left.displayName.localeCompare(right.displayName))
      .forEach((folder) => {
        const option = document.createElement("option");
        option.value = folder.id;
        option.textContent = folder.displayName;
        if (folder.displayName.toLowerCase() === "inbox") option.selected = true;
        fragment.append(option);
      });

    elements.folderSelect.replaceChildren(fragment);
    if (folders.length === 0) {
      throw new Error("No Outlook mail folders were returned for this account.");
    }
  }

  function messagesUrl(folderId, search) {
    const params = new URLSearchParams({
      "$top": String(MESSAGE_PAGE_SIZE),
      "$select": "id,subject,from,receivedDateTime,bodyPreview,body",
      "$orderby": "receivedDateTime desc",
    });
    if (search) {
      params.set("$search", `"${search.replace(/["\\]/g, " ").trim()}"`);
    }
    return `${GRAPH_ROOT}/me/mailFolders/${encodeURIComponent(folderId)}/messages?${params}`;
  }

  async function loadMessages({ append = false, search = "" } = {}) {
    if (state.loading) return;
    const folderId = elements.folderSelect.value;
    if (!folderId) return;

    clearStatus();
    setBusy(true, append ? elements.loadMoreButton : null);
    elements.listStatus.textContent = append ? "Loading more…" : "Loading messages…";
    if (!append) {
      state.messages = [];
      state.selectedIds.clear();
      state.nextLink = null;
      renderMessages();
    }

    try {
      const url = append ? validateNextLink(state.nextLink) : messagesUrl(folderId, search);
      const payload = await graphGet(url);
      const incoming = Array.isArray(payload.value) ? payload.value : [];
      state.messages = append ? [...state.messages, ...incoming] : incoming;
      state.nextLink = payload["@odata.nextLink"] || null;
      elements.loadMoreButton.hidden = !state.nextLink;
      renderMessages();
    } catch (error) {
      elements.listStatus.textContent = "";
      showStatus(error.message || "Could not load messages. Please try again.", "error");
    } finally {
      setBusy(false);
      renderMessages();
    }
  }

  function exportText(message) {
    const sender = message.from && message.from.emailAddress ? message.from.emailAddress : {};
    const lines = [
      `Subject: ${cleanText(message.subject) || "(No subject)"}`,
      `From: ${cleanText(sender.name)}${sender.address ? ` <${cleanText(sender.address)}>` : ""}`,
      `Received: ${formatDate(message.receivedDateTime)}`,
      "",
      normalizedBody(message) || cleanText(message.bodyPreview) || "(No message text)",
    ];
    return lines.join("\n").trim();
  }

  async function copySelected() {
    if (state.selectedIds.size === 0 || state.loading) return;
    const selected = state.messages.filter((message) => state.selectedIds.has(message.id));
    if (selected.length === 0) return;

    setBusy(true, elements.copyButton);
    clearStatus();
    try {
      const output = selected.map(exportText).join("\n\n" + "-".repeat(60) + "\n\n");
      await navigator.clipboard.writeText(output);
      showStatus(`Copied ${selected.length} ${selected.length === 1 ? "email" : "emails"} as plain text. Paste it into Copilot when you’re ready.`, "success");
    } catch (error) {
      showStatus(error.message || "Could not copy to the clipboard. Check clipboard permissions and try again.", "error");
    } finally {
      setBusy(false);
      renderMessages();
    }
  }

  async function connect() {
    if (state.loading) return;
    clearStatus();
    const clientId = elements.clientId.value.trim();
    if (!clientId) {
      showStatus("Enter your Microsoft app (client) ID first. The setup guide explains how to get one.", "error");
      elements.clientId.focus();
      return;
    }

    setBusy(true, elements.connectButton);
    try {
      localStorage.setItem(CLIENT_ID_KEY, clientId);
      state.client = makeClient(clientId);
      if (typeof state.client.initialize === "function") {
        await state.client.initialize();
      }
      const result = await state.client.loginPopup({ scopes: [MAIL_SCOPE] });
      state.account = result.account;
      state.token = result.accessToken;
      setSignedIn(result.account);
      await loadFolders();
      setBusy(false);
      await loadMessages();
    } catch (error) {
      if (
        window.msal &&
        window.msal.BrowserAuthError &&
        error instanceof window.msal.BrowserAuthError &&
        error.errorCode === "user_cancelled"
      ) {
        showStatus("Sign-in was cancelled. You can connect whenever you’re ready.");
      } else {
        showStatus(error.message || "Could not connect to Outlook. Check the app ID and sign-in setup, then try again.", "error");
      }
    } finally {
      setBusy(false);
    }
  }

  async function signOut() {
    if (!state.client || !state.account) return;
    try {
      await state.client.logoutPopup({ account: state.account });
    } catch (error) {
      showStatus(error.message || "Microsoft could not complete sign-out.", "error");
      return;
    }
    state.account = null;
    state.token = null;
    state.folders = [];
    state.messages = [];
    state.selectedIds.clear();
    state.nextLink = null;
    elements.mailPanel.hidden = true;
    elements.connectPanel.hidden = false;
    elements.signOutButton.hidden = true;
    elements.accountLabel.hidden = true;
    elements.folderSelect.replaceChildren();
    elements.messageList.replaceChildren();
    elements.listStatus.textContent = "";
    elements.searchInput.value = "";
    clearStatus();
  }

  elements.connectButton.addEventListener("click", connect);
  elements.signOutButton.addEventListener("click", signOut);
  elements.copyButton.addEventListener("click", copySelected);
  elements.refreshButton.addEventListener("click", () => loadMessages({ search: elements.searchInput.value.trim() }));
  elements.loadMoreButton.addEventListener("click", () => loadMessages({ append: true }));
  elements.folderSelect.addEventListener("change", () => {
    elements.searchInput.value = "";
    loadMessages();
  });
  elements.selectAll.addEventListener("change", () => {
    const visibleMessages = filteredMessages();
    visibleMessages.forEach((message) => {
      if (elements.selectAll.checked) state.selectedIds.add(message.id);
      else state.selectedIds.delete(message.id);
    });
    updateSelection();
  });
  elements.searchInput.addEventListener("input", () => {
    window.clearTimeout(state.searchTimer);
    state.searchTimer = window.setTimeout(() => {
      loadMessages({ search: elements.searchInput.value.trim() });
    }, 350);
  });

  try {
    elements.clientId.value = localStorage.getItem(CLIENT_ID_KEY) || DEFAULT_CLIENT_ID;
  } catch {
    elements.clientId.value = DEFAULT_CLIENT_ID;
    showStatus("Browser storage is disabled. Your app ID will need to be entered each time.", "info");
  }
})();
