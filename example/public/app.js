'use strict';

const STORAGE_KEY = 'unimail-auth';
const DEFAULT_LIMIT = 50;

const els = {
    connector: document.querySelector('#connector'),
    accessToken: document.querySelector('#access_token'),
    refreshToken: document.querySelector('#refresh_token'),
    expirationDate: document.querySelector('#expiration_date'),
    save: document.querySelector('#save'),
    status: document.querySelector('#status'),
    folders: document.querySelector('#folders'),
    messages: document.querySelector('#messages'),
    messagesTitle: document.querySelector('#messages-title'),
    messagesBanner: document.querySelector('#messages-banner'),
    loadMore: document.querySelector('#load-more'),
    detailSubject: document.querySelector('#detail-subject'),
    detailMeta: document.querySelector('#detail-meta'),
    detailBody: document.querySelector('#detail-body'),
    detailFiles: document.querySelector('#detail-files'),
    filters: document.querySelector('#filters'),
    filtersActive: document.querySelector('#filters-active'),
    filterSubject: document.querySelector('#filter-subject'),
    filterAfter: document.querySelector('#filter-after'),
    filterBefore: document.querySelector('#filter-before'),
    filterParticipants: document.querySelector('#filter-participants'),
    filterIncludeDrafts: document.querySelector('#filter-include-drafts'),
    filterApply: document.querySelector('#filter-apply'),
    filterClear: document.querySelector('#filter-clear'),
    showRaw: document.querySelector('#show-raw'),
    rawDialog: document.querySelector('#raw-dialog'),
    rawContent: document.querySelector('#raw-content'),
    rawClose: document.querySelector('#raw-close')
};

const state = {
    folders: [],
    selectedFolder: null,
    selectedMessageId: null,
    messages: [],
    nextPageToken: null,
    filters: {
        subject: '',
        after: '',
        before: '',
        participants: [],
        includeDrafts: false
    }
};

function loadAuth() {

    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) {
            return null;
        }

        return JSON.parse(raw);
    }
    catch {
        return null;
    }
}

function saveAuth(auth) {

    localStorage.setItem(STORAGE_KEY, JSON.stringify(auth));
}

function getAuthFromInputs() {

    const existing = loadAuth() || {};

    return {
        connector: els.connector.value,
        access_token: els.accessToken.value.trim(),
        refresh_token: els.refreshToken.value.trim(),
        expiration_date: els.expirationDate.value.trim() || null,
        id: existing.id || (crypto.randomUUID ? crypto.randomUUID() : `${String(Date.now())  }-${  Math.random().toString(36).slice(2)}`)
    };
}

function fillInputsFromAuth(auth) {

    if (!auth) {
        return;
    }

    if (auth.connector) {
        els.connector.value = auth.connector;
    }

    els.accessToken.value = auth.access_token || '';
    els.refreshToken.value = auth.refresh_token || '';
    els.expirationDate.value = auth.expiration_date || '';
}

function setStatus(message, kind) {

    els.status.textContent = message || '';
    els.status.className = `status${  kind ? ` status-${  kind}` : ''}`;
}

function authHeaderValue(auth) {

    const safe = {
        access_token: auth.access_token,
        refresh_token: auth.refresh_token,
        expiration_date: auth.expiration_date || null,
        id: auth.id
    };

    return btoa(unescape(encodeURIComponent(JSON.stringify(safe))));
}

function decodeAuthHeader(value) {

    return JSON.parse(decodeURIComponent(escape(atob(value))));
}

async function api(path) {

    const auth = loadAuth();
    if (!auth || !auth.connector) {
        throw new Error('Please save connector and tokens first.');
    }

    const res = await fetch(path, {
        headers: {
            'X-Unimail-Connector': auth.connector,
            'X-Unimail-Auth': authHeaderValue(auth)
        }
    });

    const updated = res.headers.get('X-Unimail-Auth-Updated');
    if (updated) {
        try {
            const newAuth = decodeAuthHeader(updated);
            const merged = { ...auth, ...newAuth, connector: auth.connector };
            saveAuth(merged);
            fillInputsFromAuth(merged);
            setStatus('Access token refreshed', 'ok');
        }
        catch { /* ignore decode errors */ }
    }

    const body = await res.json().catch(() => null);

    if (!res.ok) {
        const message = body && (body.message || body.error) || res.statusText;
        throw new Error(`${res.status}: ${message}`);
    }

    return body;
}

function clearChildren(node) {

    while (node.firstChild) {
        node.firstChild.remove();
    }
}

function renderFolders() {

    clearChildren(els.folders);

    state.folders.forEach((folder) => {

        const li = document.createElement('li');
        li.className = 'item folder-item';
        if (state.selectedFolder && folder.id === state.selectedFolder.id) {
            li.classList.add('selected');
        }

        const role = folder.role ? ` (${folder.role})` : '';
        const counts = folder.unread_count != null ? ` · ${folder.unread_count} unread` : '';
        li.innerHTML = `<span class="folder-name">${escapeHtml(folder.name || folder.id)}</span><span class="folder-meta">${escapeHtml(role + counts)}</span>`;
        li.addEventListener('click', () => selectFolder(folder));
        els.folders.append(li);
    });
}

function renderMessagesLoading() {

    clearChildren(els.messages);
    const li = document.createElement('li');
    li.className = 'item loading';
    li.innerHTML = '<span class="spinner" aria-hidden="true"></span><span>Loading messages...</span>';
    els.messages.append(li);
    els.loadMore.hidden = true;
}

function renderMessages() {

    clearChildren(els.messages);

    if (state.messages.length === 0) {
        const li = document.createElement('li');
        li.className = 'item empty';
        li.textContent = 'No messages';
        els.messages.append(li);
    }

    state.messages.forEach((message) => {

        const li = document.createElement('li');
        li.className = 'item message-item';

        const from = message.addresses && message.addresses.from;
        const fromText = from ? (from.name || from.email || '') : '';
        const date = formatDate(message.date);

        li.innerHTML = `
            <div class="message-from">${escapeHtml(fromText)}</div>
            <div class="message-subject">${escapeHtml(message.subject || '(no subject)')}</div>
            <div class="message-date">${escapeHtml(date)}</div>
        `;

        li.addEventListener('click', () => selectMessage(message));
        els.messages.append(li);
    });

    els.loadMore.hidden = !state.nextPageToken;
}

function renderMessageDetail(message) {

    els.detailSubject.textContent = message.subject || '(no subject)';

    const addresses = message.addresses || {};
    const from = addresses.from || {};
    const to = addresses.to || [];
    const cc = addresses.cc || [];

    const meta = [
        ['From', formatAddress(from)],
        ['To', to.map(formatAddress).filter(Boolean).join(', ')],
        ['Cc', cc.map(formatAddress).filter(Boolean).join(', ')],
        ['Date', formatDate(message.date)],
        ['Folders', (message.folders || []).join(', ')]
    ];

    clearChildren(els.detailMeta);
    meta.forEach(([label, value]) => {

        if (!value) {
            return;
        }

        const dt = document.createElement('dt');
        dt.textContent = label;
        const dd = document.createElement('dd');
        dd.textContent = value;
        els.detailMeta.append(dt);
        els.detailMeta.append(dd);
    });

    clearChildren(els.detailBody);
    const body = pickBody(message.body || []);
    if (body && body.type === 'text/html') {
        const iframe = document.createElement('iframe');
        iframe.sandbox = '';
        iframe.srcdoc = body.content || '';
        els.detailBody.append(iframe);
    }
    else if (body) {
        const pre = document.createElement('pre');
        pre.textContent = body.content || '';
        els.detailBody.append(pre);
    }
    else {
        els.detailBody.textContent = '(no body)';
    }

    clearChildren(els.detailFiles);
    const files = message.files || [];
    if (files.length > 0) {
        const heading = document.createElement('h3');
        heading.textContent = `Attachments (${files.length})`;
        els.detailFiles.append(heading);

        const ul = document.createElement('ul');
        ul.className = 'list files-list';

        files.forEach((file) => {

            const li = document.createElement('li');
            li.className = 'item file-item';
            const sizeKb = typeof file.size === 'number' ? `${file.size} KB` : '';
            li.innerHTML = `
                <div class="file-name">${escapeHtml(file.file_name || '(unnamed)')}</div>
                <div class="file-meta">${escapeHtml([file.type || '', sizeKb].filter(Boolean).join(' · '))}</div>
            `;
            ul.append(li);
        });

        els.detailFiles.append(ul);
    }
}

function formatAddress(address) {

    if (!address) {
        return '';
    }

    if (address.name && address.email) {
        return `${address.name} <${address.email}>`;
    }

    return address.email || address.name || '';
}

function pickBody(bodies) {

    return bodies.find((b) => b.type === 'text/html') || bodies.find((b) => b.type === 'text/plain') || bodies[0] || null;
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function formatDate(value) {

    if (value == null || value === '') {
        return '';
    }

    let date;
    if (typeof value === 'number') {
        date = new Date(value > 1e12 ? value : value * 1000);
    }
    else {
        date = new Date(value);
    }

    if (Number.isNaN(date.getTime())) {
        return '';
    }

    const month = MONTHS[date.getMonth()];
    const day = date.getDate();
    const year = String(date.getFullYear()).slice(-2);
    const rawHours = date.getHours();
    const period = rawHours >= 12 ? 'PM' : 'AM';
    const hours = rawHours % 12 || 12;
    const minutes = String(date.getMinutes()).padStart(2, '0');

    return `${month} ${day}, ${year} ${hours}:${minutes} ${period}`;
}

function escapeHtml(value) {

    return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function folderQueryParam(folder) {

    if (folder.virtualAll) {
        return null;
    }

    const auth = loadAuth();
    if (!auth) {
        return null;
    }

    if (auth.connector === 'gmail') {
        return folder.name;
    }

    if (auth.connector === 'unipile') {
        return folder.role || folder.name;
    }

    if (auth.connector === 'office365') {
        if (folder.role === 'sent') {
            return 'sent';
        }

        return null;
    }

    return folder.name;
}

function showMessagesBanner(text) {

    if (!text) {
        els.messagesBanner.hidden = true;
        els.messagesBanner.textContent = '';

        return;
    }

    els.messagesBanner.hidden = false;
    els.messagesBanner.textContent = text;
}

async function selectFolder(folder) {

    state.selectedFolder = folder;
    renderFolders();
    els.messagesTitle.textContent = folder.name || folder.id;

    const auth = loadAuth();
    const folderParam = folderQueryParam(folder);

    if (auth && auth.connector === 'office365' && !folder.virtualAll && folder.role !== 'inbox' && folder.role !== 'sent') {
        showMessagesBanner('Filtering by this folder is not supported by the Office 365 connector. Showing all messages instead.');
    }
    else {
        showMessagesBanner('');
    }

    await loadMessages(folderParam);
}

async function loadMessages(folderParam, append) {

    const query = new URLSearchParams();
    if (folderParam) {
        query.set('folder', folderParam);
    }

    query.set('limit', String(DEFAULT_LIMIT));

    if (append && state.nextPageToken) {
        query.set('pageToken', state.nextPageToken);
    }

    const filters = state.filters;
    if (filters.subject) {
        query.set('subject', filters.subject);
    }

    if (filters.after) {
        query.set('after', filters.after);
    }

    if (filters.before) {
        query.set('before', filters.before);
    }

    filters.participants.forEach((email) => query.append('participants', email));

    if (filters.includeDrafts) {
        query.set('includeDrafts', 'true');
    }

    if (!append) {
        state.messages = [];
        state.nextPageToken = null;
        renderMessagesLoading();
    }

    setStatus('Loading messages...');

    try {
        const data = await api(`/api/messages?${query.toString()}`);
        const messages = Array.isArray(data) ? data : (data.messages || []);
        state.messages = append ? [...state.messages, ...messages] : messages;
        state.nextPageToken = (data && data.next_page_token) || null;
        renderMessages();
        setStatus('');
    }
    catch (err) {
        setStatus(`Failed to load messages: ${err.message}`, 'error');
    }
}

async function selectMessage(message) {

    const id = message.service_message_id || message.id;
    state.selectedMessageId = id;
    els.showRaw.hidden = false;
    els.detailSubject.textContent = 'Loading...';
    clearChildren(els.detailMeta);
    clearChildren(els.detailBody);
    clearChildren(els.detailFiles);

    try {
        const full = await api(`/api/messages/${encodeURIComponent(id)}`);
        renderMessageDetail(full);
    }
    catch (err) {
        els.detailSubject.textContent = 'Error';
        const pre = document.createElement('pre');
        pre.textContent = err.message;
        els.detailBody.append(pre);
    }
}

async function showRawMessage() {

    if (!state.selectedMessageId) {
        return;
    }

    els.rawContent.textContent = 'Loading...';
    if (typeof els.rawDialog.showModal === 'function') {
        els.rawDialog.showModal();
    }
    else {
        els.rawDialog.setAttribute('open', '');
    }

    try {
        const raw = await api(`/api/messages/${encodeURIComponent(state.selectedMessageId)}?raw=true`);
        els.rawContent.textContent = typeof raw === 'string' ? raw : JSON.stringify(raw, null, 2);
    }
    catch (err) {
        els.rawContent.textContent = `Failed to load raw message: ${err.message}`;
    }
}

function closeRawDialog() {

    if (typeof els.rawDialog.close === 'function') {
        els.rawDialog.close();
    }
    else {
        els.rawDialog.removeAttribute('open');
    }
}

async function loadFolders() {

    setStatus('Loading folders...');

    try {
        const data = await api('/api/folders');
        const fetched = Array.isArray(data) ? data : [];
        const virtualAll = {
            id: '__all__',
            name: 'All emails',
            role: null,
            virtualAll: true
        };
        state.folders = [virtualAll, ...fetched];
        renderFolders();

        const inbox = fetched.find((f) => f.role === 'inbox');
        if (inbox) {
            await selectFolder(inbox);
        }
        else if (fetched.length > 0) {
            await selectFolder(fetched[0]);
        }
        else {
            await selectFolder(virtualAll);
        }
    }
    catch (err) {
        setStatus(`Failed to load folders: ${err.message}`, 'error');
    }
}

function readFiltersFromInputs() {

    const participants = els.filterParticipants.value
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter(Boolean);

    return {
        subject: els.filterSubject.value.trim(),
        after: els.filterAfter.value,
        before: els.filterBefore.value,
        participants,
        includeDrafts: els.filterIncludeDrafts.checked
    };
}

function countActiveFilters(filters) {

    let count = 0;
    if (filters.subject) {
        count++;
    }

    if (filters.after) {
        count++;
    }

    if (filters.before) {
        count++;
    }

    if (filters.participants.length > 0) {
        count++;
    }

    if (filters.includeDrafts) {
        count++;
    }

    return count;
}

function updateFiltersBadge() {

    const n = countActiveFilters(state.filters);
    if (n === 0) {
        els.filtersActive.hidden = true;
        els.filtersActive.textContent = '';
    }
    else {
        els.filtersActive.hidden = false;
        els.filtersActive.textContent = String(n);
    }
}

function resetFilterInputs() {

    els.filterSubject.value = '';
    els.filterAfter.value = '';
    els.filterBefore.value = '';
    els.filterParticipants.value = '';
    els.filterIncludeDrafts.checked = false;
}

async function populateConnectors() {

    try {
        const res = await fetch('/connectors');
        const list = await res.json();
        clearChildren(els.connector);
        list.forEach((name) => {

            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            els.connector.append(opt);
        });

        if (list.length === 0) {
            setStatus('No connectors configured. Set env vars in example/.env and restart the server.', 'error');
        }
    }
    catch (err) {
        setStatus(`Failed to load connectors: ${err.message}`, 'error');
    }
}

function init() {

    els.save.addEventListener('click', async () => {

        const auth = getAuthFromInputs();
        if (!auth.access_token) {
            setStatus('Access token is required', 'error');

            return;
        }

        saveAuth(auth);
        setStatus('Saved');
        await loadFolders();
    });

    els.loadMore.addEventListener('click', () => {

        if (!state.selectedFolder) {
            return;
        }

        loadMessages(folderQueryParam(state.selectedFolder), true);
    });

    els.showRaw.addEventListener('click', () => showRawMessage());
    els.rawClose.addEventListener('click', () => closeRawDialog());
    els.rawDialog.addEventListener('click', (event) => {

        if (event.target === els.rawDialog) {
            closeRawDialog();
        }
    });

    els.filterApply.addEventListener('click', (event) => {

        event.preventDefault();
        state.filters = readFiltersFromInputs();
        updateFiltersBadge();

        if (state.selectedFolder) {
            loadMessages(folderQueryParam(state.selectedFolder));
        }
    });

    els.filterClear.addEventListener('click', () => {

        resetFilterInputs();
        state.filters = readFiltersFromInputs();
        updateFiltersBadge();

        if (state.selectedFolder) {
            loadMessages(folderQueryParam(state.selectedFolder));
        }
    });

    populateConnectors().then(() => {

        const existing = loadAuth();
        if (existing) {
            fillInputsFromAuth(existing);
            if (existing.access_token) {
                loadFolders();
            }
        }
    });
}

init();
