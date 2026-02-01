// Email AI Manager Dashboard JavaScript
// Comprehensive feature-rich version

const API_BASE = '/api';

// State
let currentTab = 'dashboard';
let currentPage = 1;
let selectedEmails = new Set();
let theme = localStorage.getItem('theme') || 'light';

// ============================================================
// Initialization
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    setupNavigation();
    setupKeyboardShortcuts();
    setupSearch();
    setupStatCardClicks();
    checkBotStatus();
    loadDashboard();

    // Periodic updates
    setInterval(checkBotStatus, 30000);
    setInterval(() => {
        if (currentTab === 'dashboard') loadDashboardStats();
    }, 30000);
});

function setupStatCardClicks() {
    // Set up click handlers for dashboard stat cards
    document.getElementById('stat-total')?.addEventListener('click', () => viewEmailsByState(''));
    document.getElementById('stat-pending')?.addEventListener('click', () => switchTab('pending'));
    document.getElementById('stat-sent')?.addEventListener('click', () => viewEmailsByState('sent'));
    document.getElementById('stat-spam')?.addEventListener('click', () => viewEmailsByState('spam_detected'));
}

// ============================================================
// Theme Management
// ============================================================

function initTheme() {
    document.documentElement.setAttribute('data-theme', theme);
    updateThemeIcon();
}

function toggleTheme() {
    theme = theme === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
    updateThemeIcon();
    showToast('Theme changed to ' + theme + ' mode', 'info');
}

function updateThemeIcon() {
    const icon = document.getElementById('theme-icon');
    icon.textContent = theme === 'light' ? '🌙' : '☀️';
}

// ============================================================
// Navigation
// ============================================================

function setupNavigation() {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => {
            const tab = item.dataset.tab;
            switchTab(tab);
        });
    });
}

function switchTab(tabName) {
    // Update nav active state
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.toggle('active', item.dataset.tab === tabName);
    });

    // Update tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`${tabName}-tab`).classList.add('active');

    // Update page title
    const titles = {
        'dashboard': 'Dashboard',
        'emails': 'All Emails',
        'pending': 'Pending Action',
        'analytics': 'Analytics',
        'audit': 'Audit Log',
        'rules': 'Spam Rules',
        'email-rules': 'Email Rules',
        'folders': 'Mailbox Folders',
        'muted': 'Muted Senders',
        'settings': 'Settings'
    };
    document.getElementById('page-title').textContent = titles[tabName] || tabName;

    currentTab = tabName;

    // Load content for the tab
    loadTabContent(tabName);
}

function loadTabContent(tabName) {
    switch (tabName) {
        case 'dashboard':
            loadDashboard();
            break;
        case 'emails':
            loadEmails();
            break;
        case 'pending':
            loadPendingEmails();
            break;
        case 'analytics':
            loadAnalytics();
            break;
        case 'audit':
            loadAuditLog();
            break;
        case 'rules':
            loadSpamRules();
            break;
        case 'email-rules':
            loadEmailRules();
            break;
        case 'folders':
            loadFoldersTab();
            break;
        case 'muted':
            loadMutedSenders();
            break;
        case 'settings':
            loadSettings();
            break;
    }
}

function refreshCurrentTab() {
    loadTabContent(currentTab);
    showToast('Refreshed', 'success');
}

function viewEmailsByState(state) {
    // Switch to emails tab and set the state filter
    switchTab('emails');
    document.getElementById('state-filter').value = state;
    document.getElementById('category-filter').value = '';
    currentPage = 1;
    loadEmails();
}

function toggleSidebar() {
    document.querySelector('.sidebar').classList.toggle('collapsed');
}

// ============================================================
// Dashboard
// ============================================================

async function loadDashboard() {
    await Promise.all([
        loadDashboardStats(),
        loadDashboardPending(),
        loadDashboardActivity(),
        loadDashboardSenders()
    ]);
}

async function loadDashboardStats() {
    try {
        const response = await fetch(`${API_BASE}/stats?hours=24`);
        const stats = await response.json();

        document.getElementById('dash-total').textContent = stats.total_emails || 0;
        document.getElementById('dash-pending').textContent = stats.pending_count || 0;
        document.getElementById('dash-sent').textContent = stats.emails_sent || 0;
        document.getElementById('dash-spam').textContent = stats.spam_filtered || 0;

        // Update nav badges
        const totalBadge = document.getElementById('nav-badge-total');
        const pendingBadge = document.getElementById('nav-badge-pending');

        if (stats.total_emails > 0) {
            totalBadge.textContent = stats.total_emails;
            totalBadge.style.display = 'inline-block';
        } else {
            totalBadge.style.display = 'none';
        }

        if (stats.pending_count > 0) {
            pendingBadge.textContent = stats.pending_count;
            pendingBadge.style.display = 'inline-block';
        } else {
            pendingBadge.style.display = 'none';
        }
    } catch (error) {
        console.error('Failed to load dashboard stats:', error);
    }
}

async function loadDashboardPending() {
    const container = document.getElementById('dash-pending-list');
    try {
        const response = await fetch(`${API_BASE}/emails/pending`);
        const data = await response.json();

        if (data.emails.length === 0) {
            container.innerHTML = '<p class="empty-state">No emails pending action</p>';
            return;
        }

        container.innerHTML = data.emails.slice(0, 5).map(email => `
            <div class="quick-email-item" onclick="showEmailDetail('${email.id}')">
                <div class="quick-email-priority priority-${email.priority}"></div>
                <div class="quick-email-content">
                    <div class="quick-email-subject">${escapeHtml(email.subject)}</div>
                    <div class="quick-email-meta">${escapeHtml(email.sender_name || email.sender_email)} - ${formatTimeAgo(email.received_at)}</div>
                </div>
                <div class="quick-email-actions" onclick="event.stopPropagation()">
                    <button class="btn btn-xs btn-success" onclick="approveEmail('${email.id}')" title="Approve & Send Reply">✓</button>
                    <button class="btn btn-xs btn-info" onclick="markFyi('${email.id}')" title="Info only - no reply needed">ℹ️</button>
                    <button class="btn btn-xs" onclick="ignoreEmail('${email.id}')" title="Ignore this time">✗</button>
                </div>
            </div>
        `).join('');
    } catch (error) {
        container.innerHTML = '<p class="empty-state">Failed to load</p>';
    }
}

async function loadDashboardActivity() {
    const container = document.getElementById('dash-activity');
    try {
        const response = await fetch(`${API_BASE}/audit?page_size=10`);
        const data = await response.json();

        if (data.entries.length === 0) {
            container.innerHTML = '<p class="empty-state">No recent activity</p>';
            return;
        }

        container.innerHTML = data.entries.slice(0, 5).map(entry => `
            <div class="activity-item">
                <span class="activity-time">${formatTimeAgo(entry.timestamp)}</span>
                <span class="activity-text">${formatActivityText(entry)}</span>
            </div>
        `).join('');
    } catch (error) {
        container.innerHTML = '<p class="empty-state">Failed to load</p>';
    }
}

async function loadDashboardSenders() {
    const container = document.getElementById('dash-senders');
    try {
        const response = await fetch(`${API_BASE}/stats/advanced?hours=24`);
        const data = await response.json();

        if (!data.top_senders || data.top_senders.length === 0) {
            container.innerHTML = '<p class="empty-state">No senders today</p>';
            return;
        }

        container.innerHTML = data.top_senders.slice(0, 5).map((sender, i) => `
            <div class="sender-item">
                <span class="sender-rank">${i + 1}</span>
                <span class="sender-email">${truncateEmail(sender.sender_email)}</span>
                <span class="sender-count">${sender.count}</span>
            </div>
        `).join('');
    } catch (error) {
        container.innerHTML = '<p class="empty-state">Failed to load</p>';
    }
}

// ============================================================
// Emails Tab
// ============================================================

async function loadEmails() {
    const listEl = document.getElementById('emails-list');
    listEl.innerHTML = '<p class="loading">Loading emails...</p>';

    try {
        const state = document.getElementById('state-filter').value;
        const category = document.getElementById('category-filter').value;
        let url = `${API_BASE}/emails?page=${currentPage}&page_size=20`;
        if (state) url += `&state=${state}`;
        if (category) url += `&category=${category}`;

        const response = await fetch(url);
        const data = await response.json();

        renderEmailList(listEl, data.emails, false);
        updatePagination(data);
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Failed to load emails</p>';
        console.error('Failed to load emails:', error);
    }
}

async function loadPendingEmails() {
    const listEl = document.getElementById('pending-list');
    listEl.innerHTML = '<p class="loading">Loading pending emails...</p>';

    try {
        const response = await fetch(`${API_BASE}/emails/pending`);
        const data = await response.json();

        if (data.emails.length === 0) {
            listEl.innerHTML = '<p class="empty-state">No emails pending action</p>';
        } else {
            renderEmailList(listEl, data.emails, true);
        }
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Failed to load pending emails</p>';
        console.error('Failed to load pending emails:', error);
    }
}

function renderEmailList(container, emails, showActions = false) {
    if (emails.length === 0) {
        container.innerHTML = '<p class="empty-state">No emails found</p>';
        return;
    }

    container.innerHTML = emails.map(email => `
        <div class="email-item ${selectedEmails.has(email.id) ? 'selected' : ''}" data-id="${email.id}">
            <div class="email-checkbox" onclick="event.stopPropagation(); toggleEmailSelection('${email.id}')">
                <input type="checkbox" ${selectedEmails.has(email.id) ? 'checked' : ''}>
            </div>
            <div class="email-priority priority-${email.priority}"></div>
            <div class="email-content" onclick="showEmailDetail('${email.id}')">
                <div class="email-subject">${escapeHtml(email.subject)}</div>
                <div class="email-meta">
                    <span class="email-sender">${escapeHtml(email.sender_name || email.sender_email)}</span>
                    <span class="email-time">${formatDate(email.received_at)}</span>
                </div>
            </div>
            <div class="email-badges">
                <span class="badge badge-state badge-${email.state.replace(/_/g, '-')}">${formatState(email.state)}</span>
                ${email.category === 'spam_candidate' ?
                    `<span class="badge badge-spam-candidate clickable-badge" onclick="event.stopPropagation(); confirmSpam('${email.id}')" title="Click to confirm as spam">🗑️ Spam?</span>` :
                    (email.category ? `<span class="badge badge-category">${formatCategory(email.category)}</span>` : '')}
                ${email.has_draft ? '<span class="badge badge-draft">Has Draft</span>' : ''}
            </div>
            ${!['sent', 'archived', 'ignored', 'spam_detected'].includes(email.state) ? `
                <div class="email-actions" onclick="event.stopPropagation()">
                    ${showActions && email.state === 'awaiting_approval' && email.has_draft ? `<button class="btn btn-success btn-sm" onclick="approveEmail('${email.id}')">Approve</button>` : ''}
                    <button class="btn btn-outline btn-sm" onclick="dismissEmail('${email.id}')" title="Hide from dashboard (keeps in mailbox)">✓</button>
                    <button class="btn btn-warning btn-sm" onclick="deleteEmail('${email.id}')" title="Delete from mailbox">🗑</button>
                    <button class="btn btn-danger btn-sm" onclick="markSpam('${email.id}')" title="Block sender">🚫</button>
                </div>
            ` : ''}
        </div>
    `).join('');
}

function updatePagination(data) {
    document.getElementById('page-info').textContent = `Page ${data.page}`;
    document.getElementById('prev-page').disabled = data.page <= 1;
    document.getElementById('next-page').disabled = !data.has_more;
}

function prevPage() {
    if (currentPage > 1) {
        currentPage--;
        loadEmails();
    }
}

function nextPage() {
    currentPage++;
    loadEmails();
}

// ============================================================
// Email Selection & Bulk Actions
// ============================================================

function toggleEmailSelection(emailId) {
    if (selectedEmails.has(emailId)) {
        selectedEmails.delete(emailId);
    } else {
        selectedEmails.add(emailId);
    }
    updateBulkActionsUI();

    // Update visual state
    const emailEl = document.querySelector(`.email-item[data-id="${emailId}"]`);
    if (emailEl) {
        emailEl.classList.toggle('selected', selectedEmails.has(emailId));
        const checkbox = emailEl.querySelector('input[type="checkbox"]');
        if (checkbox) checkbox.checked = selectedEmails.has(emailId);
    }
}

function updateBulkActionsUI() {
    const bulkActions = document.getElementById('bulk-actions');
    const selectedCount = document.getElementById('selected-count');

    if (selectedEmails.size > 0) {
        bulkActions.style.display = 'flex';
        selectedCount.textContent = `${selectedEmails.size} selected`;
    } else {
        bulkActions.style.display = 'none';
    }
}

function clearSelection() {
    selectedEmails.clear();
    updateBulkActionsUI();
    document.querySelectorAll('.email-item.selected').forEach(el => {
        el.classList.remove('selected');
        const checkbox = el.querySelector('input[type="checkbox"]');
        if (checkbox) checkbox.checked = false;
    });
}

async function bulkApprove() {
    if (selectedEmails.size === 0) return;

    const confirmed = confirm(`Approve and send ${selectedEmails.size} email(s)?`);
    if (!confirmed) return;

    showToast(`Approving ${selectedEmails.size} emails...`, 'info');

    let success = 0;
    let failed = 0;

    for (const emailId of selectedEmails) {
        try {
            const response = await fetch(`${API_BASE}/emails/${emailId}/approve`, { method: 'POST' });
            const result = await response.json();
            if (result.success) success++;
            else failed++;
        } catch {
            failed++;
        }
    }

    clearSelection();
    loadTabContent(currentTab);
    loadDashboardStats();

    if (failed === 0) {
        showToast(`Successfully approved ${success} email(s)`, 'success');
    } else {
        showToast(`Approved ${success}, failed ${failed}`, 'warning');
    }
}

async function bulkIgnore() {
    if (selectedEmails.size === 0) return;

    showToast(`Ignoring ${selectedEmails.size} emails...`, 'info');

    let success = 0;

    for (const emailId of selectedEmails) {
        try {
            const response = await fetch(`${API_BASE}/emails/${emailId}/ignore`, { method: 'POST' });
            if (response.ok) success++;
        } catch { /* ignore */ }
    }

    clearSelection();
    loadTabContent(currentTab);
    loadDashboardStats();
    showToast(`Ignored ${success} email(s)`, 'success');
}

async function bulkDelete() {
    if (selectedEmails.size === 0) return;

    const confirmed = confirm(`Delete ${selectedEmails.size} email(s)? (Does not block senders)`);
    if (!confirmed) return;

    showToast(`Deleting ${selectedEmails.size} emails...`, 'info');

    let success = 0;

    for (const emailId of selectedEmails) {
        try {
            const response = await fetch(`${API_BASE}/emails/${emailId}/delete`, { method: 'POST' });
            if (response.ok) success++;
        } catch { /* ignore */ }
    }

    clearSelection();
    loadTabContent(currentTab);
    loadDashboardStats();
    showToast(`Deleted ${success} email(s)`, 'success');
}

async function bulkSpam() {
    if (selectedEmails.size === 0) return;

    showToast(`Marking ${selectedEmails.size} as spam...`, 'info');

    let success = 0;

    for (const emailId of selectedEmails) {
        try {
            const response = await fetch(`${API_BASE}/emails/${emailId}/spam`, { method: 'POST' });
            if (response.ok) success++;
        } catch { /* ignore */ }
    }

    clearSelection();
    loadTabContent(currentTab);
    loadDashboardStats();
    showToast(`Marked ${success} email(s) as spam`, 'success');
}

async function bulkDismiss() {
    if (selectedEmails.size === 0) return;

    showToast(`Dismissing ${selectedEmails.size} emails from dashboard...`, 'info');

    let success = 0;

    for (const emailId of selectedEmails) {
        try {
            const response = await fetch(`${API_BASE}/emails/${emailId}/dismiss`, { method: 'POST' });
            if (response.ok) success++;
        } catch { /* ignore */ }
    }

    clearSelection();
    loadTabContent(currentTab);
    loadDashboardStats();
    showToast(`Dismissed ${success} email(s) from dashboard`, 'success');
}

// ============================================================
// Email Detail Modal
// ============================================================

async function showEmailDetail(emailId) {
    const modal = document.getElementById('email-modal');
    const detailEl = document.getElementById('email-detail');

    detailEl.innerHTML = '<p class="loading">Loading email details...</p>';
    modal.classList.add('open');

    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}`);
        const email = await response.json();

        document.getElementById('modal-title').textContent = email.subject || 'Email Details';

        detailEl.innerHTML = `
            <div class="email-detail-header">
                <div class="detail-row">
                    <span class="detail-label">From:</span>
                    <span class="detail-value">${escapeHtml(email.sender_name || '')} &lt;${escapeHtml(email.sender_email)}&gt;</span>
                    <button class="btn btn-xs btn-outline" onclick="muteSenderQuick('${escapeHtml(email.sender_email)}')" title="Mute this sender">🔇</button>
                </div>
                <div class="detail-row">
                    <span class="detail-label">To:</span>
                    <span class="detail-value">${(email.to_recipients || []).map(escapeHtml).join(', ')}</span>
                </div>
                ${email.cc_recipients && email.cc_recipients.length > 0 ? `
                    <div class="detail-row">
                        <span class="detail-label">CC:</span>
                        <span class="detail-value">${email.cc_recipients.map(escapeHtml).join(', ')}</span>
                    </div>
                ` : ''}
                <div class="detail-row">
                    <span class="detail-label">Received:</span>
                    <span class="detail-value">${formatFullDate(email.received_at)}</span>
                </div>
                <div class="detail-meta">
                    <span class="badge badge-state badge-${email.state.replace(/_/g, '-')}">${formatState(email.state)}</span>
                    ${email.category ? `<span class="badge badge-category">${formatCategory(email.category)}</span>` : ''}
                    <span class="badge">Priority: ${email.priority}/5</span>
                    ${email.spam_score > 0 ? `<span class="badge badge-spam">Spam: ${email.spam_score}%</span>` : ''}
                </div>
            </div>

            ${email.summary ? `
                <div class="email-summary">
                    <h4>AI Summary</h4>
                    <p>${escapeHtml(email.summary)}</p>
                </div>
            ` : ''}

            <div class="email-body-container">
                <h4>Email Content</h4>
                <div class="email-body">${formatEmailBody(email.body_full || email.body_preview)}</div>
            </div>

            ${email.current_draft ? `
                <div class="draft-section">
                    <h4>Draft Reply</h4>
                    <div class="draft-content">${escapeHtml(email.current_draft)}</div>
                </div>
            ` : ''}

            <div class="edit-draft-section">
                <h4>${email.current_draft ? 'Edit Draft' : 'Create Reply'}</h4>
                <textarea id="draft-instructions" placeholder="Type your instructions here... e.g. 'Make it shorter' or 'Say we can do Thursday instead' or 'Tell them yes but ask about pricing'" rows="3" style="width:100%;padding:10px;border:1px solid var(--border);border-radius:4px;font-family:inherit;resize:vertical;"></textarea>
                <div style="margin-top:8px;">
                    <button class="btn btn-primary" onclick="regenerateDraft('${email.id}')">${email.current_draft ? '✨ Update Draft' : '✨ Generate Draft'}</button>
                </div>
            </div>

            <div class="modal-actions">
                ${email.state === 'awaiting_approval' && email.current_draft ? `
                    <button class="btn btn-success" onclick="approveEmail('${email.id}')">Approve & Send</button>
                ` : ''}
                ${!['sent', 'archived', 'ignored', 'spam_detected'].includes(email.state) ? `
                    ${!['fyi_notified', 'acknowledged'].includes(email.state) ? `
                        <button class="btn btn-info" onclick="markFyi('${email.id}')">Mark as Info</button>
                    ` : ''}
                    <button class="btn btn-outline" onclick="dismissEmail('${email.id}')" title="Hide from dashboard (keeps in mailbox)">Dismiss</button>
                    <button class="btn btn-warning" onclick="deleteEmail('${email.id}')" title="Delete from mailbox">Delete</button>
                    <button class="btn btn-danger" onclick="markSpam('${email.id}')" title="Delete & block sender">Spam</button>
                ` : ''}
                <button class="btn btn-outline" onclick="closeModal()">Close</button>
            </div>
        `;
    } catch (error) {
        detailEl.innerHTML = '<p class="empty-state">Failed to load email details</p>';
        console.error('Failed to load email detail:', error);
    }
}

function closeModal() {
    document.getElementById('email-modal').classList.remove('open');
}

function closeShortcutsModal() {
    document.getElementById('shortcuts-modal').classList.remove('open');
}

function showShortcutsModal() {
    document.getElementById('shortcuts-modal').classList.add('open');
}

// ============================================================
// Email Actions
// ============================================================

async function approveEmail(emailId) {
    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/approve`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Email approved for sending', 'success');
            closeModal();
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function ignoreEmail(emailId) {
    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/ignore`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Email ignored', 'success');
            closeModal();
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

// Confirm dialog promise resolver
let confirmResolver = null;

function showConfirmDialog(title, message, actionText = 'Confirm', buttonClass = 'btn-danger') {
    return new Promise((resolve) => {
        confirmResolver = resolve;
        document.getElementById('confirm-title').textContent = title;
        document.getElementById('confirm-message').textContent = message;
        const actionBtn = document.getElementById('confirm-action-btn');
        actionBtn.textContent = actionText;
        actionBtn.className = `btn ${buttonClass}`;
        document.getElementById('confirm-modal').classList.add('open');
    });
}

function closeConfirmDialog(result) {
    document.getElementById('confirm-modal').classList.remove('open');
    if (confirmResolver) {
        confirmResolver(result);
        confirmResolver = null;
    }
}

async function dismissEmail(emailId) {
    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/dismiss`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Email dismissed from dashboard', 'success');
            closeModal();
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function deleteEmail(emailId) {
    const confirmed = await showConfirmDialog(
        'Delete Email',
        'Delete this email from your mailbox? This moves it to Deleted Items in MS365.',
        'Yes, Delete',
        'btn-warning'
    );

    if (!confirmed) return;

    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/delete`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Email deleted from mailbox', 'success');
            closeModal();
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function regenerateDraft(emailId) {
    const instructions = document.getElementById('draft-instructions').value.trim();

    if (!instructions) {
        showToast('Please enter instructions for the draft', 'warning');
        return;
    }

    showToast('Generating draft...', 'info');

    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/regenerate-draft`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ instructions: instructions })
        });
        const result = await response.json();

        if (result.success) {
            showToast('Draft updated!', 'success');
            // Refresh the modal to show new draft
            showEmailDetail(emailId);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function markSpam(emailId) {
    const confirmed = await showConfirmDialog(
        'Mark as Spam',
        'Mark this email as spam? This will delete it AND block future emails from this sender.',
        'Yes, Block Sender',
        'btn-danger'
    );

    if (!confirmed) return;

    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/spam`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Marked as spam and sender blocked', 'success');
            closeModal();
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function markFyi(emailId) {
    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/fyi`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Marked as FYI', 'success');
            closeModal();
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function confirmSpam(emailId) {
    // Quick confirm spam from the badge - marks as spam and learns the pattern
    const confirmed = await showConfirmDialog(
        'Confirm Spam',
        'Confirm this is spam? It will be deleted and the pattern will be learned.',
        'Yes, It\'s Spam',
        'btn-danger'
    );

    if (!confirmed) return;

    try {
        const response = await fetch(`${API_BASE}/emails/${emailId}/spam`, { method: 'POST' });
        const result = await response.json();

        if (result.success) {
            showToast('Confirmed as spam - pattern learned', 'success');
            loadDashboardStats();
            loadTabContent(currentTab);
        } else {
            showToast('Failed: ' + result.message, 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

// ============================================================
// Audit Log
// ============================================================

async function loadAuditLog() {
    const listEl = document.getElementById('audit-list');
    listEl.innerHTML = '<p class="loading">Loading audit log...</p>';

    try {
        const agent = document.getElementById('audit-agent-filter').value;
        let url = `${API_BASE}/audit?page_size=100`;
        if (agent) url += `&agent=${agent}`;

        const response = await fetch(url);
        const data = await response.json();

        if (data.entries.length === 0) {
            listEl.innerHTML = '<p class="empty-state">No audit entries</p>';
            return;
        }

        listEl.innerHTML = data.entries.map(entry => `
            <div class="audit-item ${entry.error ? 'has-error' : ''}">
                <div class="audit-header">
                    <span class="audit-time">${formatFullDate(entry.timestamp)}</span>
                    <span class="audit-agent badge">${entry.agent}</span>
                    <span class="audit-action">${entry.action}</span>
                </div>
                ${entry.user_command ? `<div class="audit-command">Command: ${escapeHtml(entry.user_command)}</div>` : ''}
                ${entry.details ? `<div class="audit-details">${formatAuditDetails(entry.details)}</div>` : ''}
                ${entry.error ? `<div class="audit-error">Error: ${escapeHtml(entry.error)}</div>` : ''}
            </div>
        `).join('');
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Failed to load audit log</p>';
        console.error('Failed to load audit log:', error);
    }
}

function formatAuditDetails(details) {
    if (typeof details === 'string') return escapeHtml(details);
    if (typeof details === 'object') {
        return Object.entries(details)
            .map(([key, value]) => `<span class="detail-item">${key}: ${escapeHtml(String(value))}</span>`)
            .join(' ');
    }
    return '';
}

// ============================================================
// Spam Rules
// ============================================================

async function loadSpamRules() {
    const listEl = document.getElementById('rules-list');
    listEl.innerHTML = '<p class="loading">Loading spam rules...</p>';

    try {
        const response = await fetch(`${API_BASE}/spam-rules`);
        const data = await response.json();

        if (data.rules.length === 0) {
            listEl.innerHTML = '<p class="empty-state">No spam rules configured. Add rules above to automatically filter unwanted emails.</p>';
            return;
        }

        listEl.innerHTML = data.rules.map(rule => `
            <div class="rule-item">
                <div class="rule-main">
                    <span class="badge badge-type">${rule.rule_type}</span>
                    <span class="rule-pattern">${escapeHtml(rule.pattern)}</span>
                    <span class="badge badge-action">${rule.action || 'archive'}</span>
                </div>
                <div class="rule-stats">
                    <span title="Times matched">Hits: ${rule.hit_count}</span>
                    <span title="False positives">FP: ${rule.false_positives}</span>
                    <span class="badge badge-confidence">Confidence: ${rule.confidence}%</span>
                </div>
                <button class="btn btn-xs btn-danger" onclick="deleteSpamRule('${rule.id}')" title="Delete rule">×</button>
            </div>
        `).join('');
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Failed to load spam rules</p>';
        console.error('Failed to load spam rules:', error);
    }
}

async function addSpamRule() {
    const ruleType = document.getElementById('rule-type').value;
    const pattern = document.getElementById('rule-pattern').value.trim();
    const action = document.getElementById('rule-action').value;
    const confidence = parseInt(document.getElementById('rule-confidence').value) || 80;

    if (!pattern) {
        showToast('Please enter a pattern', 'warning');
        return;
    }

    try {
        const response = await fetch(`${API_BASE}/spam-rules`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                rule_type: ruleType,
                pattern: pattern,
                action: action,
                confidence: confidence
            })
        });

        if (response.ok) {
            document.getElementById('rule-pattern').value = '';
            loadSpamRules();
            showToast('Spam rule added', 'success');
        } else {
            const error = await response.json();
            showToast('Failed: ' + (error.detail || 'Unknown error'), 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function deleteSpamRule(ruleId) {
    if (!confirm('Delete this spam rule?')) return;

    try {
        const response = await fetch(`${API_BASE}/spam-rules/${ruleId}`, { method: 'DELETE' });
        if (response.ok) {
            loadSpamRules();
            showToast('Spam rule deleted', 'success');
        } else {
            showToast('Failed to delete rule', 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

// ============================================================
// Muted Senders
// ============================================================

async function loadMutedSenders() {
    const listEl = document.getElementById('muted-list');
    listEl.innerHTML = '<p class="loading">Loading muted senders...</p>';

    try {
        const response = await fetch(`${API_BASE}/muted-senders`);
        const data = await response.json();

        if (!data.senders || data.senders.length === 0) {
            listEl.innerHTML = '<p class="empty-state">No muted senders. Muted senders won\'t trigger Teams notifications.</p>';
            return;
        }

        listEl.innerHTML = data.senders.map(sender => `
            <div class="muted-item">
                <div class="muted-main">
                    <span class="muted-pattern">${escapeHtml(sender.pattern)}</span>
                    ${sender.reason ? `<span class="muted-reason">${escapeHtml(sender.reason)}</span>` : ''}
                </div>
                <div class="muted-meta">
                    <span class="muted-date">Muted: ${formatDate(sender.muted_at)}</span>
                </div>
                <button class="btn btn-xs btn-outline" onclick="unmuteSender('${escapeHtml(sender.pattern)}')" title="Unmute">Unmute</button>
            </div>
        `).join('');
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Failed to load muted senders</p>';
        console.error('Failed to load muted senders:', error);
    }
}

async function muteSender() {
    const pattern = document.getElementById('mute-pattern').value.trim();
    const reason = document.getElementById('mute-reason').value.trim();

    if (!pattern) {
        showToast('Please enter an email or domain', 'warning');
        return;
    }

    try {
        const response = await fetch(`${API_BASE}/muted-senders`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pattern, reason })
        });

        if (response.ok) {
            document.getElementById('mute-pattern').value = '';
            document.getElementById('mute-reason').value = '';
            loadMutedSenders();
            showToast('Sender muted', 'success');
        } else {
            const error = await response.json();
            showToast('Failed: ' + (error.detail || 'Unknown error'), 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function muteSenderQuick(email) {
    const reason = 'Muted from email detail';
    try {
        const response = await fetch(`${API_BASE}/muted-senders`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pattern: email, reason })
        });

        if (response.ok) {
            showToast(`Muted ${email}`, 'success');
        } else {
            showToast('Failed to mute sender', 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function unmuteSender(pattern) {
    try {
        const response = await fetch(`${API_BASE}/muted-senders/${encodeURIComponent(pattern)}`, {
            method: 'DELETE'
        });

        if (response.ok) {
            loadMutedSenders();
            showToast('Sender unmuted', 'success');
        } else {
            showToast('Failed to unmute sender', 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

// ============================================================
// Analytics
// ============================================================

async function loadAnalytics() {
    const hours = document.getElementById('analytics-hours').value;

    try {
        const response = await fetch(`${API_BASE}/stats/advanced?hours=${hours}`);
        const data = await response.json();

        renderAnalyticsSummary(data);
        renderCategoryChart(data.by_category);
        renderPriorityChart(data.by_priority);
        renderMailboxChart(data.by_mailbox);
        renderTopSenders(data.top_senders);
        renderHourlyChart(data.hourly_distribution);
    } catch (error) {
        console.error('Failed to load analytics:', error);
        showToast('Failed to load analytics', 'error');
    }
}

function renderAnalyticsSummary(data) {
    const total = Object.values(data.by_category || {}).reduce((a, b) => a + b, 0);

    document.getElementById('analytics-total').textContent = total;
    document.getElementById('analytics-auto-sent').textContent = data.auto_sent || 0;
    document.getElementById('analytics-vip').textContent = data.vip_emails || 0;
    document.getElementById('analytics-meetings').textContent = data.meeting_emails || 0;
    document.getElementById('analytics-avg-response').textContent =
        data.avg_response_minutes ? Math.round(data.avg_response_minutes) : '-';
}

function renderCategoryChart(categoryData) {
    const container = document.getElementById('category-chart');

    if (!categoryData || Object.keys(categoryData).length === 0) {
        container.innerHTML = '<p class="empty-state">No data available</p>';
        return;
    }

    const maxValue = Math.max(...Object.values(categoryData));

    container.innerHTML = Object.entries(categoryData)
        .sort((a, b) => b[1] - a[1])
        .map(([category, count]) => `
            <div class="bar-row">
                <span class="bar-label">${formatCategory(category)}</span>
                <div class="bar-container">
                    <div class="bar" style="width: ${(count / maxValue) * 100}%"></div>
                </div>
                <span class="bar-value">${count}</span>
            </div>
        `).join('');
}

function renderPriorityChart(priorityData) {
    const container = document.getElementById('priority-chart');

    if (!priorityData || Object.keys(priorityData).length === 0) {
        container.innerHTML = '<p class="empty-state">No data available</p>';
        return;
    }

    const maxValue = Math.max(...Object.values(priorityData));
    const priorityLabels = { '1': 'Critical', '2': 'High', '3': 'Normal', '4': 'Low', '5': 'Minimal' };
    const priorityColors = { '1': 'var(--danger)', '2': 'var(--warning)', '3': 'var(--primary)', '4': 'var(--text-muted)', '5': 'var(--border)' };

    container.innerHTML = Object.entries(priorityData)
        .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
        .map(([priority, count]) => `
            <div class="bar-row">
                <span class="bar-label">${priorityLabels[priority] || priority}</span>
                <div class="bar-container">
                    <div class="bar" style="width: ${(count / maxValue) * 100}%; background: ${priorityColors[priority] || 'var(--primary)'}"></div>
                </div>
                <span class="bar-value">${count}</span>
            </div>
        `).join('');
}

function renderMailboxChart(mailboxData) {
    const container = document.getElementById('mailbox-chart');

    if (!mailboxData || Object.keys(mailboxData).length === 0) {
        container.innerHTML = '<p class="empty-state">No data available</p>';
        return;
    }

    const maxValue = Math.max(...Object.values(mailboxData));

    container.innerHTML = Object.entries(mailboxData)
        .sort((a, b) => b[1] - a[1])
        .map(([mailbox, count]) => `
            <div class="bar-row">
                <span class="bar-label" title="${mailbox}">${truncateEmail(mailbox)}</span>
                <div class="bar-container">
                    <div class="bar bar-mailbox" style="width: ${(count / maxValue) * 100}%"></div>
                </div>
                <span class="bar-value">${count}</span>
            </div>
        `).join('');
}

function renderTopSenders(sendersData) {
    const container = document.getElementById('top-senders');

    if (!sendersData || sendersData.length === 0) {
        container.innerHTML = '<p class="empty-state">No data available</p>';
        return;
    }

    container.innerHTML = sendersData.map((sender, index) => `
        <div class="sender-row">
            <span class="sender-rank">${index + 1}</span>
            <span class="sender-email" title="${sender.sender_email}">${truncateEmail(sender.sender_email)}</span>
            <span class="sender-count">${sender.count} emails</span>
        </div>
    `).join('');
}

function renderHourlyChart(hourlyData) {
    const container = document.getElementById('hourly-chart');

    if (!hourlyData || Object.keys(hourlyData).length === 0) {
        container.innerHTML = '<p class="empty-state">No data available</p>';
        return;
    }

    const allHours = {};
    for (let i = 0; i < 24; i++) {
        allHours[i.toString()] = hourlyData[i.toString()] || 0;
    }

    const maxValue = Math.max(...Object.values(allHours), 1);

    container.innerHTML = `
        <div class="hourly-bars">
            ${Object.entries(allHours).map(([hour, count]) => `
                <div class="hourly-bar-container" title="${hour}:00 - ${count} emails">
                    <div class="hourly-bar" style="height: ${(count / maxValue) * 100}%"></div>
                </div>
            `).join('')}
        </div>
        <div class="hourly-labels">
            <span>0</span>
            <span>6</span>
            <span>12</span>
            <span>18</span>
            <span>24</span>
        </div>
    `;
}

// ============================================================
// Settings
// ============================================================

async function loadSettings() {
    try {
        const response = await fetch(`${API_BASE}/settings`);
        if (response.ok) {
            const settings = await response.json();
            populateSettings(settings);
        }
    } catch (error) {
        console.error('Failed to load settings:', error);
    }

    // Check MCP status
    checkMCPStatus();
}

function populateSettings(settings) {
    if (settings.poll_interval_seconds) {
        document.getElementById('setting-poll-interval').value = settings.poll_interval_seconds;
    }
    if (settings.mailbox_email) {
        document.getElementById('setting-mailbox').value = settings.mailbox_email;
    }
    if (settings.teams_morning_summary_hour) {
        document.getElementById('setting-morning-hour').value = settings.teams_morning_summary_hour;
    }
    if (settings.teams_notify_urgent !== undefined) {
        document.getElementById('setting-notify-urgent').checked = settings.teams_notify_urgent;
    }
    if (settings.max_emails_per_hour) {
        document.getElementById('setting-max-emails').value = settings.max_emails_per_hour;
    }
    if (settings.agent_model) {
        document.getElementById('ai-model').textContent = settings.agent_model;
    }
    if (settings.db_path) {
        document.getElementById('db-path').textContent = settings.db_path;
    }
}

async function checkMCPStatus() {
    const statusEl = document.getElementById('mcp-status');
    try {
        const response = await fetch(`${API_BASE}/health`);
        const data = await response.json();

        if (data.mcp_connected) {
            statusEl.innerHTML = '<span class="status-dot online"></span><span>Connected</span>';
        } else {
            statusEl.innerHTML = '<span class="status-dot offline"></span><span>Disconnected</span>';
        }

        if (data.teams_configured) {
            document.getElementById('teams-channel-status').textContent = 'Configured';
        } else {
            document.getElementById('teams-channel-status').textContent = 'Not configured';
        }
    } catch (error) {
        statusEl.innerHTML = '<span class="status-dot offline"></span><span>Error</span>';
    }
}

// ============================================================
// Bot Status
// ============================================================

async function checkBotStatus() {
    const statusEl = document.getElementById('bot-status');
    const syncEl = document.getElementById('last-sync');

    try {
        const response = await fetch(`${API_BASE}/health`);
        const data = await response.json();

        if (data.status === 'healthy') {
            statusEl.innerHTML = '<span class="status-dot online"></span><span class="status-text">Running</span>';
        } else {
            statusEl.innerHTML = '<span class="status-dot offline"></span><span class="status-text">Unhealthy</span>';
        }

        if (data.last_poll) {
            syncEl.textContent = `Last sync: ${formatTimeAgo(data.last_poll)}`;
        }

        if (data.uptime_seconds) {
            document.getElementById('uptime').textContent = formatUptime(data.uptime_seconds);
        }
    } catch (error) {
        statusEl.innerHTML = '<span class="status-dot offline"></span><span class="status-text">Offline</span>';
        syncEl.textContent = 'Last sync: --';
    }
}

// ============================================================
// Search
// ============================================================

function setupSearch() {
    const searchInput = document.getElementById('global-search');
    let debounceTimeout;

    searchInput.addEventListener('input', (e) => {
        clearTimeout(debounceTimeout);
        debounceTimeout = setTimeout(() => {
            const query = e.target.value.trim();
            if (query.length >= 2) {
                searchEmails(query);
            } else if (query.length === 0 && currentTab === 'emails') {
                loadEmails();
            }
        }, 300);
    });

    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            searchInput.value = '';
            searchInput.blur();
            if (currentTab === 'emails') loadEmails();
        }
    });
}

async function searchEmails(query) {
    // Switch to emails tab if not there
    if (currentTab !== 'emails') {
        switchTab('emails');
    }

    const listEl = document.getElementById('emails-list');
    listEl.innerHTML = '<p class="loading">Searching...</p>';

    try {
        const response = await fetch(`${API_BASE}/emails?search=${encodeURIComponent(query)}&page_size=50`);
        const data = await response.json();

        renderEmailList(listEl, data.emails, false);
        document.getElementById('page-info').textContent = `${data.emails.length} results`;
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Search failed</p>';
    }
}

// ============================================================
// Keyboard Shortcuts
// ============================================================

function setupKeyboardShortcuts() {
    document.addEventListener('keydown', (e) => {
        // Don't handle shortcuts when typing in inputs
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
            return;
        }

        // Ctrl+K - Focus search
        if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
            e.preventDefault();
            document.getElementById('global-search').focus();
            return;
        }

        // Escape - Close modal
        if (e.key === 'Escape') {
            closeModal();
            closeShortcutsModal();
            return;
        }

        // R - Refresh
        if (e.key === 'r' || e.key === 'R') {
            e.preventDefault();
            refreshCurrentTab();
            return;
        }

        // ? - Show shortcuts
        if (e.key === '?') {
            e.preventDefault();
            showShortcutsModal();
            return;
        }

        // Number keys - Switch tabs
        const tabMap = { '1': 'dashboard', '2': 'emails', '3': 'pending', '4': 'analytics', '5': 'audit', '6': 'rules', '7': 'muted', '8': 'settings' };
        if (tabMap[e.key]) {
            e.preventDefault();
            switchTab(tabMap[e.key]);
            return;
        }
    });
}

// ============================================================
// Toast Notifications
// ============================================================

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const icons = { success: '✓', error: '✗', warning: '⚠', info: 'ℹ' };
    toast.innerHTML = `
        <span class="toast-icon">${icons[type] || icons.info}</span>
        <span class="toast-message">${escapeHtml(message)}</span>
    `;

    container.appendChild(toast);

    // Trigger animation
    setTimeout(() => toast.classList.add('show'), 10);

    // Auto-remove after 4 seconds
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// ============================================================
// Utility Functions
// ============================================================

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatDate(dateStr) {
    if (!dateStr) return '--';
    const date = new Date(dateStr);
    const now = new Date();
    const diff = now - date;

    if (diff < 86400000) {
        return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }
    if (diff < 604800000) {
        return date.toLocaleDateString([], { weekday: 'short', hour: '2-digit', minute: '2-digit' });
    }
    return date.toLocaleDateString([], { month: 'short', day: 'numeric' });
}

function formatFullDate(dateStr) {
    if (!dateStr) return '--';
    const date = new Date(dateStr);
    return date.toLocaleString([], {
        weekday: 'short',
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function formatTimeAgo(dateStr) {
    if (!dateStr) return '--';
    const date = new Date(dateStr);
    const now = new Date();
    const diff = Math.floor((now - date) / 1000);

    if (diff < 60) return 'just now';
    if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
    if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
    if (diff < 604800) return `${Math.floor(diff / 86400)}d ago`;
    return formatDate(dateStr);
}

function formatUptime(seconds) {
    if (!seconds) return '--';
    const days = Math.floor(seconds / 86400);
    const hours = Math.floor((seconds % 86400) / 3600);
    const mins = Math.floor((seconds % 3600) / 60);

    if (days > 0) return `${days}d ${hours}h`;
    if (hours > 0) return `${hours}h ${mins}m`;
    return `${mins}m`;
}

function formatState(state) {
    const states = {
        'new': 'New',
        'processing': 'Processing',
        'spam_detected': 'Spam',
        'fyi_notified': 'FYI',
        'action_required': 'Action Required',
        'draft_generated': 'Draft Ready',
        'awaiting_approval': 'Awaiting Approval',
        'approved': 'Approved',
        'sent': 'Sent',
        'ignored': 'Ignored',
        'forward_suggested': 'Forward',
        'forwarded': 'Forwarded',
        'archived': 'Archived',
        'error': 'Error',
        'acknowledged': 'Acknowledged',
        'held_for_morning': 'Held for Morning'
    };
    return states[state] || state;
}

function formatCategory(category) {
    if (!category) return 'Unknown';
    return category.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

function formatActivityText(entry) {
    const actions = {
        'email_received': 'New email received',
        'draft_generated': 'Draft generated',
        'email_approved': 'Email approved',
        'email_sent': 'Email sent',
        'email_ignored': 'Email ignored',
        'marked_spam': 'Marked as spam',
        'notification_sent': 'Teams notification sent'
    };
    return actions[entry.action] || entry.action;
}

function formatEmailBody(body) {
    if (!body) return '<p class="empty-state">No content</p>';

    // Check if it's HTML content
    if (body.trim().startsWith('<') && (body.includes('<html') || body.includes('<body') || body.includes('<div') || body.includes('<p>') || body.includes('<table'))) {
        // It's HTML - render directly in a container with scoped styles
        return `<div class="email-html-content" style="background:white;padding:15px;border-radius:4px;border:1px solid var(--border);overflow-x:auto;">${body}</div>`;
    }

    // Plain text - escape and format
    return '<pre class="email-plaintext" style="white-space:pre-wrap;word-wrap:break-word;font-family:inherit;background:white;padding:15px;border-radius:4px;border:1px solid var(--border);">' + escapeHtml(body) + '</pre>';
}

function truncateEmail(email) {
    if (!email) return '';
    if (email.length <= 25) return email;
    const [local, domain] = email.split('@');
    if (domain && domain.length > 15) {
        return `${local.substring(0, 10)}...@${domain.substring(0, 12)}...`;
    }
    return email.substring(0, 22) + '...';
}

// ============================================================
// Email Rules (LLM-based routing)
// ============================================================

async function loadEmailRules() {
    const listEl = document.getElementById('email-rules-list');
    listEl.innerHTML = '<p class="loading">Loading email rules...</p>';

    try {
        const includeInactive = document.getElementById('show-inactive-rules')?.checked || false;
        const response = await fetch(`${API_BASE}/email-rules?include_inactive=${includeInactive}`);
        const data = await response.json();

        if (data.rules.length === 0) {
            listEl.innerHTML = `
                <div class="empty-state">
                    <p>No email rules configured.</p>
                    <p style="margin-top: 10px; color: var(--text-muted);">
                        Create a rule above to automatically route emails based on natural language conditions.
                    </p>
                </div>
            `;
            return;
        }

        listEl.innerHTML = data.rules.map(rule => `
            <div class="email-rule-item ${!rule.is_active ? 'inactive' : ''}">
                <div class="email-rule-header">
                    <div class="email-rule-name">
                        ${rule.is_active ? '' : '<span class="badge badge-muted">Inactive</span>'}
                        <strong>${escapeHtml(rule.name)}</strong>
                        <span class="badge badge-priority">Priority: ${rule.priority}</span>
                    </div>
                    <div class="email-rule-actions">
                        ${rule.is_active ? `
                            <button class="btn btn-xs btn-primary" onclick="runEmailRule('${rule.id}', true)" title="Test run (preview only)">🔍 Test</button>
                            <button class="btn btn-xs btn-success" onclick="runEmailRule('${rule.id}', false)" title="Run now on recent emails">▶️ Run</button>
                        ` : ''}
                        <button class="btn btn-xs btn-outline" onclick="toggleEmailRule('${rule.id}', ${!rule.is_active})" title="${rule.is_active ? 'Disable' : 'Enable'}">
                            ${rule.is_active ? '⏸️' : '▶️'}
                        </button>
                        <button class="btn btn-xs btn-outline" onclick="editEmailRule('${rule.id}')" title="Edit">✏️</button>
                        <button class="btn btn-xs btn-danger" onclick="deleteEmailRule('${rule.id}')" title="Delete">×</button>
                    </div>
                </div>
                <div class="email-rule-prompt">
                    <span class="label">Match:</span>
                    <span class="value">${escapeHtml(rule.match_prompt)}</span>
                </div>
                <div class="email-rule-action-row">
                    <span class="label">Action:</span>
                    <span class="badge badge-action">${formatRuleAction(rule.action)}</span>
                    ${rule.action_value ? `<span class="value">${escapeHtml(rule.action_value)}</span>` : ''}
                    ${rule.action === 'move_to_folder' && rule.action_value ? `
                        <button class="btn btn-xs btn-outline" onclick="viewFolderContents('${escapeHtml(rule.action_value)}')" title="View emails in this folder">
                            👁️ View
                        </button>
                    ` : ''}
                </div>
                <div class="email-rule-stats">
                    <span title="Times matched">Hits: ${rule.hit_count}</span>
                    ${rule.last_hit ? `<span title="Last matched">Last: ${formatTimeAgo(rule.last_hit)}</span>` : ''}
                    ${rule.false_positives > 0 ? `<span class="warning" title="False positives">FP: ${rule.false_positives}</span>` : ''}
                    ${rule.stop_processing ? '<span class="badge badge-stop" title="Stops further rule processing">Stop</span>' : ''}
                </div>
                ${rule.description ? `<div class="email-rule-description">${escapeHtml(rule.description)}</div>` : ''}
            </div>
        `).join('');
    } catch (error) {
        listEl.innerHTML = '<p class="empty-state">Failed to load email rules</p>';
        console.error('Failed to load email rules:', error);
    }
}

function formatRuleAction(action) {
    const actions = {
        'move_to_folder': '📁 Move to Folder',
        'archive': '📦 Archive',
        'forward': '↗️ Forward',
        'set_priority': '⚡ Set Priority',
        'add_label': '🏷️ Add Label',
        'notify': '🔔 Custom Notify'
    };
    return actions[action] || action;
}

async function addEmailRule() {
    const name = document.getElementById('email-rule-name').value.trim();
    const matchPrompt = document.getElementById('email-rule-prompt').value.trim();
    const action = document.getElementById('email-rule-action').value;
    const actionValue = document.getElementById('email-rule-action-value').value.trim();
    const priority = parseInt(document.getElementById('email-rule-priority').value) || 50;

    if (!name) {
        showToast('Please enter a rule name', 'warning');
        return;
    }

    if (!matchPrompt) {
        showToast('Please enter a match condition', 'warning');
        return;
    }

    if ((action === 'move_to_folder' || action === 'forward') && !actionValue) {
        showToast(`Please enter a ${action === 'move_to_folder' ? 'folder name' : 'forward address'}`, 'warning');
        return;
    }

    showToast('Creating rule...', 'info');

    try {
        const response = await fetch(`${API_BASE}/email-rules`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                name: name,
                match_prompt: matchPrompt,
                action: action,
                action_value: actionValue,
                priority: priority
            })
        });

        if (response.ok) {
            // Clear form
            document.getElementById('email-rule-name').value = '';
            document.getElementById('email-rule-prompt').value = '';
            document.getElementById('email-rule-action-value').value = '';
            document.getElementById('email-rule-priority').value = '50';

            loadEmailRules();
            showToast('Email rule created!', 'success');
        } else {
            const error = await response.json();
            showToast('Failed: ' + (error.detail || 'Unknown error'), 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function testEmailRulePrompt() {
    const matchPrompt = document.getElementById('email-rule-prompt').value.trim();

    if (!matchPrompt) {
        showToast('Please enter a match condition to test', 'warning');
        return;
    }

    const resultsEl = document.getElementById('email-rule-test-results');
    resultsEl.style.display = 'block';
    resultsEl.innerHTML = '<p class="loading">Testing rule against recent emails...</p>';

    try {
        const response = await fetch(`${API_BASE}/email-rules/test`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                match_prompt: matchPrompt,
                limit: 20
            })
        });

        const data = await response.json();

        if (data.matches_found === 0) {
            resultsEl.innerHTML = `
                <div class="test-results-header">
                    <h4>Test Results</h4>
                    <button class="btn btn-xs btn-outline" onclick="document.getElementById('email-rule-test-results').style.display='none'">×</button>
                </div>
                <p class="empty-state">No matches found in ${data.total_tested} recent emails.</p>
                <p style="color: var(--text-muted); font-size: 12px; margin-top: 10px;">
                    Try adjusting your match condition to be broader, or check that you have recent emails that would match.
                </p>
            `;
            return;
        }

        resultsEl.innerHTML = `
            <div class="test-results-header">
                <h4>Test Results</h4>
                <button class="btn btn-xs btn-outline" onclick="document.getElementById('email-rule-test-results').style.display='none'">×</button>
            </div>
            <p class="test-summary">
                Found <strong>${data.matches_found}</strong> matching emails out of ${data.total_tested} tested
            </p>
            <div class="test-matches">
                ${data.matches.map(match => `
                    <div class="test-match-item">
                        <div class="test-match-main">
                            <span class="test-match-subject">${escapeHtml(match.subject)}</span>
                            <span class="test-match-sender">${escapeHtml(match.sender)}</span>
                        </div>
                        <div class="test-match-meta">
                            <span class="badge badge-confidence" title="Confidence">
                                ${match.confidence}%
                            </span>
                            <span class="test-match-reason">${escapeHtml(match.reason)}</span>
                        </div>
                    </div>
                `).join('')}
            </div>
        `;
    } catch (error) {
        resultsEl.innerHTML = `
            <div class="test-results-header">
                <h4>Test Results</h4>
                <button class="btn btn-xs btn-outline" onclick="document.getElementById('email-rule-test-results').style.display='none'">×</button>
            </div>
            <p class="empty-state">Test failed: ${error.message}</p>
        `;
    }
}

async function toggleEmailRule(ruleId, activate) {
    try {
        const response = await fetch(`${API_BASE}/email-rules/${ruleId}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ is_active: activate })
        });

        if (response.ok) {
            loadEmailRules();
            showToast(`Rule ${activate ? 'enabled' : 'disabled'}`, 'success');
        } else {
            showToast('Failed to update rule', 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function deleteEmailRule(ruleId) {
    const confirmed = await showConfirmDialog(
        'Delete Email Rule',
        'Are you sure you want to delete this rule? This cannot be undone.',
        'Delete Rule',
        'btn-danger'
    );

    if (!confirmed) return;

    try {
        const response = await fetch(`${API_BASE}/email-rules/${ruleId}`, { method: 'DELETE' });
        if (response.ok) {
            loadEmailRules();
            showToast('Email rule deleted', 'success');
        } else {
            showToast('Failed to delete rule', 'error');
        }
    } catch (error) {
        showToast('Error: ' + error.message, 'error');
    }
}

async function editEmailRule(ruleId) {
    try {
        const response = await fetch(`${API_BASE}/email-rules/${ruleId}`);
        const rule = await response.json();

        // Populate form with rule data
        document.getElementById('email-rule-name').value = rule.name;
        document.getElementById('email-rule-prompt').value = rule.match_prompt;
        document.getElementById('email-rule-action').value = rule.action;
        document.getElementById('email-rule-action-value').value = rule.action_value || '';
        document.getElementById('email-rule-priority').value = rule.priority;

        // Scroll to form
        document.getElementById('email-rules-tab').scrollTo({ top: 0, behavior: 'smooth' });

        showToast('Rule loaded into form. Make changes and click Create to save as a new rule, or delete the old one first.', 'info');
    } catch (error) {
        showToast('Error loading rule: ' + error.message, 'error');
    }
}

function showProgressModal(title, message) {
    document.getElementById('progress-title').textContent = title;
    document.getElementById('progress-message').textContent = message;
    document.getElementById('progress-details').innerHTML = '';
    document.getElementById('progress-modal').classList.add('open');
}

function updateProgressModal(message, details = null) {
    document.getElementById('progress-message').textContent = message;
    if (details) {
        document.getElementById('progress-details').innerHTML = details;
    }
}

function closeProgressModal() {
    document.getElementById('progress-modal').classList.remove('open');
}

async function runEmailRule(ruleId, dryRun = true) {
    const actionText = dryRun ? 'Testing' : 'Running';

    if (!dryRun) {
        const confirmed = await showConfirmDialog(
            'Run Rule Now',
            'This will apply the rule action to all matching emails. Are you sure?',
            'Yes, Run Now',
            'btn-success'
        );

        if (!confirmed) return;
    }

    // Show progress modal
    showProgressModal(
        `${actionText} Rule...`,
        'Starting...'
    );

    try {
        // Use Server-Sent Events for streaming progress
        const eventSource = new EventSource(
            `${API_BASE}/email-rules/${ruleId}/run-stream?dry_run=${dryRun}&limit=50`
        );

        let matchedEmails = [];
        let result = null;

        eventSource.onmessage = (event) => {
            const data = JSON.parse(event.data);

            switch (data.type) {
                case 'start':
                    updateProgressModal(
                        `Starting: ${data.total} emails to check`,
                        `<div class="progress-info">Rule: ${escapeHtml(data.rule_name)}</div>`
                    );
                    break;

                case 'progress':
                    updateProgressModal(
                        `Checking email ${data.current} of ${data.total}...`,
                        `<div class="progress-item"><span class="subject">${escapeHtml(data.subject)}...</span></div>`
                    );
                    break;

                case 'action':
                    // Show when an action is applied
                    const detailsEl = document.getElementById('progress-details');
                    detailsEl.innerHTML += `
                        <div class="progress-item match">
                            ✓ ${escapeHtml(data.subject)}... → ${data.action}
                        </div>
                    `;
                    break;

                case 'complete':
                    result = data;
                    eventSource.close();
                    closeProgressModal();

                    if (data.matched === 0) {
                        showToast(`No matching emails found (${data.total_evaluated} checked)`, 'info');
                    } else if (dryRun) {
                        showRuleRunResults(data);
                    } else {
                        showToast(
                            `Rule applied! ${data.processed} emails processed, ${data.errors} errors`,
                            data.errors > 0 ? 'warning' : 'success'
                        );
                        if (data.matched > 0) {
                            showRuleRunResults(data);
                        }
                    }
                    loadEmailRules();
                    break;

                case 'error':
                    eventSource.close();
                    closeProgressModal();
                    showToast('Error: ' + data.message, 'error');
                    break;
            }
        };

        eventSource.onerror = () => {
            eventSource.close();
            closeProgressModal();
            showToast('Connection lost during processing', 'error');
        };

    } catch (error) {
        closeProgressModal();
        showToast('Error: ' + error.message, 'error');
    }
}

// Fallback non-streaming version (kept for compatibility)
async function runEmailRuleLegacy(ruleId, dryRun = true) {
    const actionText = dryRun ? 'Testing' : 'Running';

    if (!dryRun) {
        const confirmed = await showConfirmDialog(
            'Run Rule Now',
            'This will apply the rule action to all matching emails. Are you sure?',
            'Yes, Run Now',
            'btn-success'
        );

        if (!confirmed) return;
    }

    showProgressModal(
        `${actionText} Rule...`,
        'Evaluating emails with AI...'
    );

    try {
        const response = await fetch(`${API_BASE}/email-rules/${ruleId}/run`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                dry_run: dryRun,
                limit: 50
            })
        });

        const result = await response.json();

        closeProgressModal();

        if (!response.ok) {
            showToast('Failed: ' + (result.detail || 'Unknown error'), 'error');
            return;
        }

        // Show results in a modal or detailed toast
        if (result.matched === 0) {
            showToast(`No matching emails found (${result.total_evaluated} checked)`, 'info');
        } else if (dryRun) {
            showRuleRunResults(result);
        } else {
            showToast(
                `Rule applied! ${result.processed} emails processed, ${result.errors} errors`,
                result.errors > 0 ? 'warning' : 'success'
            );
            if (result.matched > 0) {
                showRuleRunResults(result);
            }
        }

        // Refresh rule stats
        loadEmailRules();

    } catch (error) {
        closeProgressModal();
        showToast('Error: ' + error.message, 'error');
    }
}

function showRuleRunResults(result) {
    const modal = document.getElementById('email-modal');
    const detailEl = document.getElementById('email-detail');

    document.getElementById('modal-title').textContent = `Rule Run Results: ${result.rule_name}`;

    const matchList = result.matches.map(match => `
        <div class="rule-run-match-item">
            <div class="match-subject">${escapeHtml(match.subject)}</div>
            <div class="match-meta">
                <span class="match-sender">${escapeHtml(match.sender)}</span>
                <span class="badge badge-confidence">${match.confidence}%</span>
            </div>
            <div class="match-reason">${escapeHtml(match.reason)}</div>
        </div>
    `).join('');

    detailEl.innerHTML = `
        <div class="rule-run-results">
            <div class="rule-run-summary">
                <div class="stat-item">
                    <span class="stat-value">${result.total_evaluated}</span>
                    <span class="stat-label">Emails Checked</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${result.matched}</span>
                    <span class="stat-label">Matches Found</span>
                </div>
                ${!result.dry_run ? `
                    <div class="stat-item">
                        <span class="stat-value">${result.processed}</span>
                        <span class="stat-label">Processed</span>
                    </div>
                    <div class="stat-item ${result.errors > 0 ? 'has-errors' : ''}">
                        <span class="stat-value">${result.errors}</span>
                        <span class="stat-label">Errors</span>
                    </div>
                ` : ''}
            </div>

            ${result.dry_run ? '<p class="dry-run-notice">🔍 This was a test run - no actions were taken.</p>' : ''}

            <h4>Matching Emails</h4>
            <div class="rule-run-matches">
                ${matchList || '<p class="empty-state">No matches</p>'}
            </div>

            ${result.error_details && result.error_details.length > 0 ? `
                <h4>Errors</h4>
                <div class="rule-run-errors">
                    ${result.error_details.map(err => `
                        <div class="error-item">
                            <span class="error-email">${err.email_id}</span>
                            <span class="error-msg">${escapeHtml(err.error)}</span>
                        </div>
                    `).join('')}
                </div>
            ` : ''}
        </div>

        <div class="modal-actions">
            ${result.dry_run && result.matched > 0 ? `
                <button class="btn btn-success" onclick="closeModal(); runEmailRule('${result.rule_id}', false)">
                    ▶️ Run for Real
                </button>
            ` : ''}
            <button class="btn btn-outline" onclick="closeModal()">Close</button>
        </div>
    `;

    modal.classList.add('open');
}

function updateActionValuePlaceholder() {
    const action = document.getElementById('email-rule-action').value;
    const valueInput = document.getElementById('email-rule-action-value');
    const browseBtn = document.getElementById('browse-folders-btn');

    const placeholders = {
        'move_to_folder': 'Folder name...',
        'archive': '(not required)',
        'notify': 'Custom notification message...',
        'set_priority': 'Priority (1-5)...',
        'forward': 'Email address to forward to...',
        'add_label': 'Label/category name...'
    };

    valueInput.placeholder = placeholders[action] || 'Value...';

    // Show/hide folder browse button based on action
    if (browseBtn) {
        browseBtn.style.display = action === 'move_to_folder' ? 'inline-flex' : 'none';
    }
}

async function showFolderBrowser() {
    const modal = document.getElementById('folder-modal');
    const listEl = document.getElementById('folder-list');

    modal.classList.add('open');
    listEl.innerHTML = '<p class="loading">Loading folders from mailbox...</p>';

    try {
        const response = await fetch(`${API_BASE}/email-folders?recursive=true`);
        const data = await response.json();

        if (!data.folders || data.folders.length === 0) {
            listEl.innerHTML = '<p class="empty-state">No folders found in mailbox</p>';
            return;
        }

        listEl.innerHTML = `
            <p style="font-size: 12px; color: var(--text-muted); margin-bottom: 12px;">
                Mailbox: ${escapeHtml(data.mailbox)}
            </p>
            <div class="folder-browser-list">
                ${renderFolderTree(data.folders, 0, '')}
            </div>
        `;
    } catch (error) {
        listEl.innerHTML = `<p class="empty-state">Failed to load folders: ${error.message}</p>`;
    }
}

function renderFolderTree(folders, depth, parentPath) {
    if (!folders || folders.length === 0) return '';

    return folders.map(folder => {
        const hasChildren = folder.children && folder.children.length > 0;
        const indent = depth * 16;
        const folderPath = parentPath ? `${parentPath}/${folder.name}` : folder.name;
        const childrenId = `children-${folder.id || Math.random().toString(36).substr(2, 9)}`;

        return `
            <div class="folder-tree-item">
                <div class="folder-item" style="padding-left: ${indent + 8}px;">
                    ${hasChildren ? `
                        <span class="folder-expand" onclick="toggleFolderExpand('${childrenId}', this); event.stopPropagation();">▶</span>
                    ` : `
                        <span class="folder-expand-placeholder"></span>
                    `}
                    <span class="folder-icon" onclick="selectFolder('${escapeHtml(folderPath)}')">${hasChildren ? '📂' : '📁'}</span>
                    <span class="folder-name" onclick="selectFolder('${escapeHtml(folderPath)}')">${escapeHtml(folder.name)}</span>
                    <span class="folder-count" title="Total: ${folder.total_count}, Unread: ${folder.unread_count}">
                        ${folder.total_count > 0 ? folder.total_count : ''}
                    </span>
                </div>
                ${hasChildren ? `
                    <div id="${childrenId}" class="folder-children" style="display: none;">
                        ${renderFolderTree(folder.children, depth + 1, folderPath)}
                    </div>
                ` : ''}
            </div>
        `;
    }).join('');
}

function toggleFolderExpand(childrenId, expandBtn) {
    const childrenEl = document.getElementById(childrenId);
    if (!childrenEl) return;

    const isExpanded = childrenEl.style.display !== 'none';
    childrenEl.style.display = isExpanded ? 'none' : 'block';
    expandBtn.textContent = isExpanded ? '▶' : '▼';
    expandBtn.classList.toggle('expanded', !isExpanded);
}

function closeFolderBrowser() {
    document.getElementById('folder-modal').classList.remove('open');
}

function selectFolder(folderName) {
    document.getElementById('email-rule-action-value').value = folderName;
    closeFolderBrowser();
    showToast(`Selected folder: ${folderName}`, 'success');
}

// ============================================================
// Folder Contents Viewer
// ============================================================

async function viewFolderContents(folderName, limit = 50) {
    const modal = document.getElementById('folder-contents-modal');
    const titleEl = document.getElementById('folder-contents-title');
    const listEl = document.getElementById('folder-contents-list');

    titleEl.textContent = `📁 ${folderName}`;
    listEl.innerHTML = '<p class="loading">Loading emails from folder...</p>';
    modal.classList.add('open');

    try {
        const response = await fetch(`${API_BASE}/folder-emails/${encodeURIComponent(folderName)}?limit=${limit}`);

        if (!response.ok) {
            const error = await response.json();
            listEl.innerHTML = `<p class="empty-state">Error: ${error.detail || 'Failed to load folder'}</p>`;
            return;
        }

        const data = await response.json();

        if (data.emails.length === 0) {
            listEl.innerHTML = `
                <p class="empty-state">No emails in this folder</p>
                <p style="text-align: center; color: var(--text-muted); font-size: 12px; margin-top: 10px;">
                    Mailbox: ${escapeHtml(data.mailbox)}
                </p>
            `;
            return;
        }

        listEl.innerHTML = `
            <div class="folder-contents-header">
                <span>${data.count} emails in folder</span>
                <span class="folder-contents-mailbox">${escapeHtml(data.mailbox)}</span>
            </div>
            <div class="folder-contents-emails">
                ${data.emails.map(email => `
                    <div class="folder-email-item ${email.is_read ? '' : 'unread'}">
                        <div class="folder-email-main">
                            <div class="folder-email-subject">${escapeHtml(email.subject)}</div>
                            <div class="folder-email-preview">${escapeHtml(email.body_preview)}</div>
                        </div>
                        <div class="folder-email-meta">
                            <span class="folder-email-sender">${escapeHtml(email.sender_name || email.sender_email)}</span>
                            <span class="folder-email-date">${formatDate(email.received_at)}</span>
                            ${email.has_attachments ? '<span class="folder-email-attachment" title="Has attachments">📎</span>' : ''}
                        </div>
                    </div>
                `).join('')}
            </div>
            ${data.count >= limit ? `
                <div class="folder-contents-footer">
                    <button class="btn btn-outline btn-sm" onclick="viewFolderContents('${escapeHtml(folderName)}', ${limit + 50})">
                        Load More
                    </button>
                </div>
            ` : ''}
        `;

    } catch (error) {
        listEl.innerHTML = `<p class="empty-state">Error: ${error.message}</p>`;
    }
}

function closeFolderContents() {
    document.getElementById('folder-contents-modal').classList.remove('open');
}

// ============================================================
// Folders Tab - Browse Mailbox Folders
// ============================================================

let currentFolderPath = '';
let selectedFolderEmails = new Set();

async function loadFoldersTab() {
    const treeEl = document.getElementById('folders-tree');
    treeEl.innerHTML = '<p class="loading">Loading folders...</p>';

    try {
        // Fetch folders and email rules in parallel
        const [foldersResponse, rulesResponse] = await Promise.all([
            fetch(`${API_BASE}/email-folders?recursive=true`),
            fetch(`${API_BASE}/email-rules?include_inactive=false`)
        ]);

        const foldersData = await foldersResponse.json();
        const rulesData = await rulesResponse.json();

        if (!foldersData.folders || foldersData.folders.length === 0) {
            treeEl.innerHTML = '<p class="empty-state">No folders found</p>';
            return;
        }

        // Extract folders used by rules (deduplicated)
        const ruleFolders = new Set();
        rulesData.rules.forEach(rule => {
            if (rule.action === 'move_to_folder' && rule.action_value) {
                ruleFolders.add(rule.action_value);
            }
        });

        // Build favorites list: Inbox + rule folders
        const favorites = ['Inbox', ...Array.from(ruleFolders)];

        // Build a map of folder counts for quick lookup
        const folderCounts = {};
        const findFolderCount = (folders, path = '') => {
            folders.forEach(f => {
                const fullPath = path ? `${path}/${f.name}` : f.name;
                folderCounts[fullPath] = { unread: f.unread_count || 0, total: f.total_count || 0 };
                folderCounts[f.name] = folderCounts[fullPath]; // Also store by name only
                if (f.children) findFolderCount(f.children, fullPath);
            });
        };
        findFolderCount(foldersData.folders);

        treeEl.innerHTML = `
            <div class="folder-tree-header">
                <span class="mailbox-name">${escapeHtml(foldersData.mailbox)}</span>
            </div>

            <div class="folder-favorites">
                <div class="folder-section-title">⭐ Favorites</div>
                ${favorites.map(folder => {
                    const counts = folderCounts[folder] || { unread: 0, total: 0 };
                    return `
                        <div class="folder-nav-row favorite ${currentFolderPath === folder ? 'active' : ''}"
                             onclick="selectFolderNav('${escapeHtml(folder)}')">
                            <span class="folder-nav-icon">📁</span>
                            <span class="folder-nav-name">${escapeHtml(folder)}</span>
                            ${counts.unread > 0 ? `<span class="folder-nav-badge">${counts.unread}</span>` : ''}
                        </div>
                    `;
                }).join('')}
            </div>

            <div class="folder-all">
                <div class="folder-section-title">📂 All Folders</div>
                <div class="folder-tree-list">
                    ${renderFolderTreeNav(foldersData.folders, 0, '')}
                </div>
            </div>
        `;
    } catch (error) {
        treeEl.innerHTML = `<p class="empty-state">Error: ${error.message}</p>`;
    }
}

function renderFolderTreeNav(folders, depth, parentPath) {
    if (!folders || folders.length === 0) return '';

    return folders.map(folder => {
        const hasChildren = folder.children && folder.children.length > 0;
        const indent = depth * 12;
        const folderPath = parentPath ? `${parentPath}/${folder.name}` : folder.name;
        const childrenId = `nav-children-${folder.id || Math.random().toString(36).substr(2, 9)}`;

        return `
            <div class="folder-nav-item">
                <div class="folder-nav-row ${currentFolderPath === folderPath ? 'active' : ''}"
                     style="padding-left: ${indent + 8}px;"
                     onclick="selectFolderNav('${escapeHtml(folderPath)}')">
                    ${hasChildren ? `
                        <span class="folder-nav-expand" onclick="event.stopPropagation(); toggleFolderNav('${childrenId}', this);">▶</span>
                    ` : `
                        <span class="folder-nav-expand-placeholder"></span>
                    `}
                    <span class="folder-nav-icon">${hasChildren ? '📂' : '📁'}</span>
                    <span class="folder-nav-name">${escapeHtml(folder.name)}</span>
                    ${folder.unread_count > 0 ? `<span class="folder-nav-badge">${folder.unread_count}</span>` : ''}
                </div>
                ${hasChildren ? `
                    <div id="${childrenId}" class="folder-nav-children" style="display: none;">
                        ${renderFolderTreeNav(folder.children, depth + 1, folderPath)}
                    </div>
                ` : ''}
            </div>
        `;
    }).join('');
}

function toggleFolderNav(childrenId, expandBtn) {
    const childrenEl = document.getElementById(childrenId);
    if (!childrenEl) return;

    const isExpanded = childrenEl.style.display !== 'none';
    childrenEl.style.display = isExpanded ? 'none' : 'block';
    expandBtn.textContent = isExpanded ? '▶' : '▼';
}

async function selectFolderNav(folderPath) {
    currentFolderPath = folderPath;

    // Update active state in tree
    document.querySelectorAll('.folder-nav-row').forEach(el => el.classList.remove('active'));
    event.target.closest('.folder-nav-row')?.classList.add('active');

    const headerEl = document.getElementById('folder-view-header');
    const listEl = document.getElementById('folder-view-list');

    headerEl.innerHTML = `<h3>📁 ${escapeHtml(folderPath)}</h3>`;
    listEl.innerHTML = '<p class="loading">Loading emails...</p>';

    try {
        const response = await fetch(`${API_BASE}/folder-emails/${encodeURIComponent(folderPath)}?limit=50`);

        if (!response.ok) {
            const error = await response.json();
            listEl.innerHTML = `<p class="empty-state">Error: ${error.detail || 'Failed to load folder'}</p>`;
            return;
        }

        const data = await response.json();

        if (data.emails.length === 0) {
            listEl.innerHTML = '<p class="empty-state">No emails in this folder</p>';
            return;
        }

        selectedFolderEmails.clear();
        updateFolderBulkActions();

        listEl.innerHTML = data.emails.map(email => `
            <div class="email-item folder-email ${email.is_read ? '' : 'unread'}" data-msgid="${escapeHtml(email.id)}">
                <div class="email-checkbox" onclick="event.stopPropagation(); toggleFolderEmailSelection('${escapeHtml(email.id)}')">
                    <input type="checkbox" ${selectedFolderEmails.has(email.id) ? 'checked' : ''}>
                </div>
                <div class="email-content" onclick="toggleFolderEmailSelection('${escapeHtml(email.id)}')">
                    <div class="email-subject">${escapeHtml(email.subject)}</div>
                    <div class="email-meta">
                        <span class="email-sender">${escapeHtml(email.sender_name || email.sender_email)}</span>
                        <span class="email-time">${formatDate(email.received_at)}</span>
                        ${email.has_attachments ? '<span title="Has attachments">📎</span>' : ''}
                    </div>
                    <div class="email-preview">${escapeHtml(email.body_preview)}</div>
                </div>
            </div>
        `).join('');

    } catch (error) {
        listEl.innerHTML = `<p class="empty-state">Error: ${error.message}</p>`;
    }
}

function toggleFolderEmailSelection(messageId) {
    if (selectedFolderEmails.has(messageId)) {
        selectedFolderEmails.delete(messageId);
    } else {
        selectedFolderEmails.add(messageId);
    }

    // Update visual state
    const emailEl = document.querySelector(`.folder-email[data-msgid="${messageId}"]`);
    if (emailEl) {
        emailEl.classList.toggle('selected', selectedFolderEmails.has(messageId));
        const checkbox = emailEl.querySelector('input[type="checkbox"]');
        if (checkbox) checkbox.checked = selectedFolderEmails.has(messageId);
    }

    updateFolderBulkActions();
}

function updateFolderBulkActions() {
    const actionsEl = document.getElementById('folder-bulk-actions');
    const countEl = document.getElementById('folder-selected-count');

    if (!actionsEl) return;

    if (selectedFolderEmails.size > 0) {
        actionsEl.style.display = 'flex';
        countEl.textContent = `${selectedFolderEmails.size} selected`;
    } else {
        actionsEl.style.display = 'none';
    }
}

function clearFolderSelection() {
    selectedFolderEmails.clear();
    document.querySelectorAll('.folder-email.selected').forEach(el => {
        el.classList.remove('selected');
        const checkbox = el.querySelector('input[type="checkbox"]');
        if (checkbox) checkbox.checked = false;
    });
    updateFolderBulkActions();
}

function selectAllFolderEmails() {
    document.querySelectorAll('.folder-email').forEach(el => {
        const msgId = el.dataset.msgid;
        if (msgId) {
            selectedFolderEmails.add(msgId);
            el.classList.add('selected');
            const checkbox = el.querySelector('input[type="checkbox"]');
            if (checkbox) checkbox.checked = true;
        }
    });
    updateFolderBulkActions();
}

async function deleteFolderEmails() {
    if (selectedFolderEmails.size === 0) return;

    const confirmed = await showConfirmDialog(
        'Delete Emails',
        `Delete ${selectedFolderEmails.size} email(s) from your mailbox? This moves them to Deleted Items.`,
        'Yes, Delete',
        'btn-warning'
    );

    if (!confirmed) return;

    showToast(`Deleting ${selectedFolderEmails.size} emails...`, 'info');

    let success = 0;
    let failed = 0;

    for (const messageId of selectedFolderEmails) {
        try {
            const response = await fetch(`${API_BASE}/folder-emails/delete`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ message_id: messageId })
            });
            if (response.ok) success++;
            else failed++;
        } catch {
            failed++;
        }
    }

    selectedFolderEmails.clear();

    if (failed === 0) {
        showToast(`Deleted ${success} email(s)`, 'success');
    } else {
        showToast(`Deleted ${success}, failed ${failed}`, 'warning');
    }

    // Refresh the folder view
    if (currentFolderPath) {
        selectFolderNav(currentFolderPath);
    }
}
