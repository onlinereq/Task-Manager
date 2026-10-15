/**
 * Team Task Manager - Frontend Application
 * Fully Vanilla JS (No Libraries, Desktop-First, Apple Design Language)
 * Connected to Google Apps Script Web App Backend
 */

(function () {
  'use strict';

  // ==========================================================================
  // 1. DATA MODULE & CONFIGURATION
  // ==========================================================================
  var API_URL = (typeof CONFIG_API_URL !== 'undefined' && CONFIG_API_URL && CONFIG_API_URL !== 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE')
    ? String(CONFIG_API_URL).trim()
    : '';

  function parseDateToEpoch(isoOrDateStr) {
    if (!isoOrDateStr) return 0;
    if (isoOrDateStr instanceof Date) return isoOrDateStr.getTime();
    var str = String(isoOrDateStr).trim();
    if (!str) return 0;
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/.test(str)) {
      str = str.replace(' ', 'T');
    } else if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
      str = str + 'T00:00:00';
    }
    var t = new Date(str).getTime();
    if (isNaN(t)) {
      t = new Date(isoOrDateStr).getTime();
    }
    return isNaN(t) ? 0 : t;
  }

  function generateTaskCodeFromDate(dateVal, idHint) {
    var d = dateVal ? new Date(dateVal) : new Date();
    if (isNaN(d.getTime())) d = new Date();
    var dd = String(d.getDate()).padStart(2, '0');
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var yy = String(d.getFullYear()).slice(-2);
    var base = 'TC' + dd + mm + yy;

    var hash = 1;
    if (idHint) {
      var str = String(idHint);
      var sum = 0;
      for (var i = 0; i < str.length; i++) {
        sum = (sum * 31 + str.charCodeAt(i)) % 99;
      }
      hash = Math.abs(sum) + 1;
    }
    return base + '-' + String(hash).padStart(2, '0');
  }

  function generateNewTaskCode(existingTasks, dateVal) {
    var d = dateVal ? new Date(dateVal) : new Date();
    if (isNaN(d.getTime())) d = new Date();
    var dd = String(d.getDate()).padStart(2, '0');
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var yy = String(d.getFullYear()).slice(-2);
    var base = 'TC' + dd + mm + yy;

    var used = {};
    (existingTasks || []).forEach(function (t) {
      var c = String(t.taskCode || '').toUpperCase();
      if (c.indexOf(base) === 0) {
        var parts = c.split('-');
        if (parts.length > 1) {
          var n = parseInt(parts[1], 10);
          if (!isNaN(n)) used[n] = true;
        }
      }
    });

    var seq = 1;
    while (used[seq]) {
      seq++;
    }
    return base + '-' + String(seq).padStart(2, '0');
  }

  var DataService = {
    users: [],       // Array of { name, team, status }
    tasks: [],       // Array of { taskId, taskName, description, assignedTo, deadline, status, state, createdBy, createdAt, lastUpdatedBy, lastUpdatedAt }
    history: [],     // Array of { historyId, taskId, taskName, status, note, updatedBy, timestamp }
    statuses: ['Register', 'In Progress', 'Review', 'Done'],
    lastSync: '',
    dataVersion: 2,

    cleanAndDeduplicateHistory: function (historyArray) {
      if (!Array.isArray(historyArray)) return [];
      var seenIds = new Set();
      var seenSemantic = new Set();
      var cleaned = [];

      // Sort descending by epoch timestamp first
      var sorted = [...historyArray].sort(function (a, b) {
        var tB = parseDateToEpoch(b.timestamp);
        var tA = parseDateToEpoch(a.timestamp);
        if (tB !== tA) return tB - tA;
        return String(b.historyId || '').localeCompare(String(a.historyId || ''));
      });

      sorted.forEach(function (h) {
        if (!h) return;
        var tId = String(h.taskId || '').trim();
        if (!tId) return;

        var hId = String(h.historyId || '').trim();
        var normStatus = String(h.status || '').trim() || 'Register';
        var normNote = String(h.note || '').trim();
        var normBy = String(h.updatedBy || h.user || '').trim();
        var epoch = parseDateToEpoch(h.timestamp);
        var epochMinute = Math.floor(epoch / 60000);

        // Deduplicate exact historyId
        if (hId && seenIds.has(hId)) return;

        // Deduplicate semantic duplicates (same task, status, note, author, within 1 minute)
        var semanticKey = tId.toLowerCase() + '|' + normStatus.toLowerCase() + '|' + normNote.toLowerCase() + '|' + normBy.toLowerCase() + '|' + epochMinute;
        if (seenSemantic.has(semanticKey)) return;

        if (hId) seenIds.add(hId);
        seenSemantic.add(semanticKey);

        cleaned.push({
          historyId: hId || ('HIST-' + tId + '-' + (epoch || Date.now())),
          taskId: tId,
          taskName: String(h.taskName || '').trim(),
          status: normStatus,
          note: normNote,
          updatedBy: normBy,
          timestamp: h.timestamp || new Date().toISOString()
        });
      });

      return cleaned;
    },

    initFromStorage: function () {
      try {
        var saved = localStorage.getItem('team_tasks_local_data');
        if (saved) {
          var parsed = JSON.parse(saved);
          if (typeof parsed.dataVersion === 'number') this.dataVersion = parsed.dataVersion;
          if (Array.isArray(parsed.tasks)) this.tasks = parsed.tasks;
          if (Array.isArray(parsed.history)) {
            this.history = this.cleanAndDeduplicateHistory(parsed.history);
          }
          if (Array.isArray(parsed.statuses) && parsed.statuses.length > 0) this.statuses = parsed.statuses;
          if (Array.isArray(parsed.users)) this.users = parsed.users;
          if (parsed.lastSync) this.lastSync = parsed.lastSync;
        }
      } catch (e) {}
    },

    saveToStorage: function () {
      try {
        var snapshot = {
          dataVersion: this.dataVersion,
          tasks: this.tasks,
          history: this.cleanAndDeduplicateHistory(this.history),
          statuses: this.statuses,
          users: this.users,
          lastSync: this.lastSync
        };
        localStorage.setItem('team_tasks_local_data', JSON.stringify(snapshot));
      } catch (e) {}
    },

    loadSyncData: function (payload) {
      if (!payload) return;

      // 1. Data Version check: if server has an updated data version, clear local cache completely on full sync
      if (typeof payload.dataVersion === 'number' && payload.dataVersion !== this.dataVersion) {
        if (!payload.delta) {
          this.tasks = [];
          this.history = [];
          this.lastSync = '';
        }
        this.dataVersion = payload.dataVersion;
      }

      if (payload.delta && this.tasks.length > 0) {
        // Delta sync merge
        if (Array.isArray(payload.tasks)) {
          payload.tasks.forEach(function (incoming) {
            var inId = String(incoming.taskId || '').trim().toLowerCase();
            var idx = DataService.tasks.findIndex(function (t) {
              return String(t.taskId || '').trim().toLowerCase() === inId;
            });
            if (idx !== -1) {
              DataService.tasks[idx] = incoming;
            } else {
              DataService.tasks.unshift(incoming);
            }
          });
        }

        // Prune deleted tasks using serverTaskIds
        if (Array.isArray(payload.serverTaskIds)) {
          var validIds = new Set(payload.serverTaskIds.map(function (id) {
            return String(id || '').trim().toLowerCase();
          }));
          this.tasks = this.tasks.filter(function (t) {
            var tId = String(t.taskId || '').trim().toLowerCase();
            return validIds.has(tId) || tId.includes('-init');
          });
          this.history = this.history.filter(function (h) {
            var hTaskId = String(h.taskId || '').trim().toLowerCase();
            return validIds.has(hTaskId);
          });
        }

        if (Array.isArray(payload.history)) {
          payload.history.forEach(function (incHist) {
            var inHId = String(incHist.historyId || '').trim();
            var inTId = String(incHist.taskId || '').trim().toLowerCase();
            var inEpoch = parseDateToEpoch(incHist.timestamp);
            var inNote = String(incHist.note || '').trim().toLowerCase();
            var inStatus = String(incHist.status || '').trim().toLowerCase();

            // Match either by explicit historyId OR matching semantic update within 60s
            var hIdx = DataService.history.findIndex(function (h) {
              if (inHId && String(h.historyId || '').trim() === inHId) return true;
              var hTId = String(h.taskId || '').trim().toLowerCase();
              if (hTId === inTId && String(h.status || '').trim().toLowerCase() === inStatus && String(h.note || '').trim().toLowerCase() === inNote) {
                var hEpoch = parseDateToEpoch(h.timestamp);
                if (Math.abs(hEpoch - inEpoch) <= 60000) return true;
              }
              return false;
            });

            if (hIdx !== -1) {
              DataService.history[hIdx] = incHist;
            } else {
              DataService.history.unshift(incHist);
            }
          });
        }
      } else {
        // Full snapshot replace (clears old data on device)
        this.tasks = Array.isArray(payload.tasks) ? payload.tasks : [];
        this.history = Array.isArray(payload.history) ? payload.history : [];
      }

      // Reconstruct and deduplicate history cleanly
      this.history = this.cleanAndDeduplicateHistory(this.history);

      if (Array.isArray(payload.statuses) && payload.statuses.length > 0) {
        this.statuses = payload.statuses;
      }
      if (Array.isArray(payload.users)) {
        this.users = payload.users;
      }
      if (payload.serverTime) {
        this.lastSync = payload.serverTime;
      }

      this.saveToStorage();
    },

    getTasks: function () {
      return [...this.tasks].sort(function (a, b) {
        return parseDateToEpoch(b.createdAt) - parseDateToEpoch(a.createdAt);
      });
    },

    getTaskById: function (id) {
      if (!id) return null;
      var targetId = String(id).trim().toLowerCase();
      return this.tasks.find(function (t) {
        return String(t.taskId || '').trim().toLowerCase() === targetId;
      }) || null;
    },

    getHistoryByTaskId: function (id) {
      if (!id) return [];
      var targetId = String(id).trim().toLowerCase();
      return this.history.filter(function (h) {
        return String(h.taskId || '').trim().toLowerCase() === targetId;
      }).sort(function (a, b) {
        var tB = parseDateToEpoch(b.timestamp);
        var tA = parseDateToEpoch(a.timestamp);
        if (tB !== tA) return tB - tA;
        return String(b.historyId || '').localeCompare(String(a.historyId || ''));
      });
    }
  };

  // ==========================================================================
  // 2. API CLIENT (text/plain JSON to avoid CORS preflight)
  // ==========================================================================
  async function apiCall(action, extraPayload) {
    if (!API_URL) {
      console.error('Missing API configuration: CONFIG_API_URL is not set or invalid.');
      return {
        ok: false,
        success: false,
        error: 'missing_config',
        message: 'Missing API configuration. Please configure CONFIG_API_URL in config.js.'
      };
    }

    extraPayload = extraPayload || {};
    var currentUser = state.currentUser;
    if (!currentUser) {
      try {
        var sess = sessionStorage.getItem('team_tasks_session_user');
        if (sess) currentUser = JSON.parse(sess);
      } catch (e) {}
    }

    var payload = Object.assign({
      action: action,
      name: (currentUser && currentUser.name) ? currentUser.name : (extraPayload.name || ''),
      pin: (currentUser && currentUser.pin) ? currentUser.pin : (extraPayload.pin || '')
    }, extraPayload);

    try {
      var controller = new AbortController();
      var timeoutId = setTimeout(function () { controller.abort(); }, 22000);

      var response = await fetch(API_URL, {
        method: 'POST',
        body: JSON.stringify(payload),
        signal: controller.signal
      });

      clearTimeout(timeoutId);
      var data = await response.json();
      return data;
    } catch (err) {
      return {
        ok: false,
        success: false,
        error: 'network_error',
        message: err.name === 'AbortError'
          ? 'Request timed out after 20 seconds. Backend took too long.'
          : 'Network error or backend unreachable. Please try again.'
      };
    }
  }

  /**
   * Checks if response is unauthorized or locked.
   * If rejected, triggers rollback, displays website modal message, and redirects to login.
   */
  function handleAuthFailureResponse(res, rollbackFn) {
    if (!res) return false;
    var isAuthFailure = (res.error === 'unauthorized' || res.error === 'locked' || (res.ok === false && (res.error === 'unauthorized' || res.error === 'locked')));
    if (isAuthFailure) {
      if (typeof rollbackFn === 'function') {
        try { rollbackFn(); } catch (e) {}
      }
      var msg = (res.error === 'locked')
        ? 'Account is temporarily locked due to 5 failed attempts (15 minutes). Please try again later.'
        : 'Session unauthorized or credentials invalid. Please sign in again.';
      showAlertModal(msg, res.error === 'locked' ? 'Account Locked' : 'Unauthorized');
      logout();
      return true;
    }
    return false;
  }

  // ==========================================================================
  // 2b. PENDING ACTION & RETRY QUEUE (Idempotent & Resilient)
  // ==========================================================================
  var PENDING_ACTIONS_STORAGE_KEY = 'team_tasks_pending_actions';

  function getPendingActions() {
    try {
      var raw = localStorage.getItem(PENDING_ACTIONS_STORAGE_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      return [];
    }
  }

  function savePendingActions(actions) {
    try {
      if (!actions || actions.length === 0) {
        localStorage.removeItem(PENDING_ACTIONS_STORAGE_KEY);
      } else {
        localStorage.setItem(PENDING_ACTIONS_STORAGE_KEY, JSON.stringify(actions));
      }
    } catch (e) {}
  }

  function addPendingAction(actionObj) {
    var list = getPendingActions();
    list.push(actionObj);
    savePendingActions(list);
  }

  function removePendingAction(requestId) {
    var list = getPendingActions().filter(function (a) {
      return a.requestId !== requestId;
    });
    savePendingActions(list);
  }

  function updatePendingActionRetry(requestId, retries) {
    var list = getPendingActions();
    for (var i = 0; i < list.length; i++) {
      if (list[i].requestId === requestId) {
        list[i].retries = retries;
        break;
      }
    }
    savePendingActions(list);
  }

  var isSyncingQueue = false;
  var inFlightBatchRequestIds = new Set();

  /**
   * Checks if the user is actively mid-edit:
   * Typing in title, description, notes, or any other input/textarea/editable field,
   * or has the task modal open, or is dragging/updating.
   */
  function isUserMidEdit() {
    if (state.isCreatingTask || state.isUpdatingStatus || state.isDragging) {
      return true;
    }
    if (el.taskModal && !el.taskModal.classList.contains('hidden')) {
      return true;
    }
    var ae = document.activeElement;
    if (ae && ae !== document.body) {
      var tag = (ae.tagName || '').toUpperCase();
      if (tag === 'INPUT' || tag === 'TEXTAREA' || ae.isContentEditable) {
        return true;
      }
    }
    return false;
  }

  /**
   * Syncs the local pending-actions queue with the server.
   * - Does nothing if queue is empty.
   * - Skips if user is actively mid-edit.
   * - Freezes current batch so new actions added mid-send are deferred to next cycle.
   */
  async function syncPendingQueue() {
    if (isSyncingQueue) return;
    if (!state.currentUser || state.startupInProgress) return;

    var allPending = getPendingActions();
    if (!allPending || allPending.length === 0) {
      return; // Queue is empty: DO NOTHING. No request sent, zero server calls.
    }

    if (isUserMidEdit()) {
      return; // User is actively typing or editing: skip this cycle and check on next tick
    }

    // Freeze in-flight batch: snapshot so new actions added mid-send go to next cycle
    var batch = allPending.slice();
    var batchRequestIds = new Set(batch.map(function (item) { return item.requestId; }));
    inFlightBatchRequestIds = batchRequestIds;
    isSyncingQueue = true;

    updateSyncStatus('syncing');

    try {
      if (batch.length === 1) {
        var single = batch[0];
        var res = await apiCall(single.action, single.payload);

        // Security check
        if (res && (res.error === 'unauthorized' || res.error === 'locked')) {
          handleAuthFailureResponse(res);
          return;
        }

        if (res && res.ok && res.success) {
          removePendingAction(single.requestId);
          updateSyncStatus('synced');
        } else if (res && res.error === 'forbidden') {
          removePendingAction(single.requestId);
          updateSyncStatus('error');
          showAlertModal(res.message || 'Authorization Denied', 'Authorization Denied');
        } else {
          // server_busy, network error, timeout: retain in queue for next cycle
          updateSyncStatus('error');
        }
      } else {
        // Send all pending actions together
        var batchPayload = {
          actions: batch.map(function (item) {
            return {
              action: item.action,
              requestId: item.requestId,
              payload: item.payload
            };
          })
        };
        var res = await apiCall('batchActions', batchPayload);

        if (res && (res.error === 'unauthorized' || res.error === 'locked')) {
          handleAuthFailureResponse(res);
          return;
        }

        if (res && res.ok && res.success) {
          // Remove ONLY the actions that were in this batch.
          // New actions added mid-send remain in queue for next cycle!
          var remaining = getPendingActions().filter(function (act) {
            return !batchRequestIds.has(act.requestId);
          });
          savePendingActions(remaining);
          updateSyncStatus('synced');
        } else {
          // server_busy, timeout, or network error: retain in queue for next cycle
          updateSyncStatus('error');
        }
      }
    } catch (e) {
      updateSyncStatus('error');
    } finally {
      inFlightBatchRequestIds.clear();
      isSyncingQueue = false;
    }
  }

  function flushPendingActions() {
    return syncPendingQueue();
  }

  /**
   * Executes a user mutation action with:
   * 1. Persistence in localStorage (tagged with unique requestId)
   * 2. In-flight batch isolation (new actions queued for next cycle if sync in progress)
   * 3. Automatic retry (up to 3 times with progressive delay) on server_busy / timeout / network_error
   * 4. Immediate rollback & rejection on genuine security failure (unauthorized, locked, forbidden)
   */
  async function executeUserActionWithRetry(options) {
    var actionName = options.action;
    var payload = options.payload || {};
    var taskName = options.taskName || 'Task';
    var rollbackFn = options.rollbackFn;
    var onSuccessFn = options.onSuccess;
    var maxRetries = (typeof options.maxRetries === 'number') ? options.maxRetries : 3;

    // Generate unique client requestId for idempotency
    var requestId = payload.requestId || ('REQ-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 7).toUpperCase());
    payload.requestId = requestId;

    var pendingItem = {
      requestId: requestId,
      action: actionName,
      payload: payload,
      taskName: taskName,
      retries: 0,
      createdAt: Date.now()
    };
    addPendingAction(pendingItem);

    // If a sync batch is currently in flight, don't mix into same batch mid-send.
    // It's saved in queue and will go out on next cycle!
    if (isSyncingQueue) {
      return;
    }

    // If user is actively typing or editing, defer send to the next 60s cycle
    if (isUserMidEdit()) {
      return;
    }

    async function attempt(retryCount) {
      updateSyncStatus('saving');
      var res = await apiCall(actionName, payload);

      // 1. Strict Security Rejection (Requirement 8):
      if (res && (res.error === 'unauthorized' || res.error === 'locked')) {
        removePendingAction(requestId);
        handleAuthFailureResponse(res, rollbackFn);
        return;
      }

      // 2. Permission failure
      if (res && res.error === 'forbidden') {
        removePendingAction(requestId);
        if (typeof rollbackFn === 'function') {
          try { rollbackFn(); } catch (e) {}
        }
        updateSyncStatus('error');
        showAlertModal(res.message || 'You do not have permission to perform this action.', 'Authorization Denied');
        return;
      }

      // 3. Success!
      if (res && res.ok && res.success) {
        removePendingAction(requestId);
        updateSyncStatus('synced');
        if (typeof onSuccessFn === 'function') {
          try { onSuccessFn(res); } catch (e) {}
        }
        return;
      }

      // 4. Transient failure (server_busy, network error, timeout)
      var isTransient = (!res || res.error === 'server_busy' || res.error === 'busy' || res.error === 'network_error' || res.ok === false);
      if (isTransient && retryCount < maxRetries) {
        var nextRetry = retryCount + 1;
        updatePendingActionRetry(requestId, nextRetry);
        updateSyncStatus('syncing');
        var delay = 800 * nextRetry;
        setTimeout(function () {
          attempt(nextRetry);
        }, delay);
        return;
      }

      // 5. Retries exhausted: retain in pending queue for next 60s cycle rather than losing it
      updateSyncStatus('error');
    }

    attempt(0);
  }

  // ==========================================================================
  // 3. UTILITIES & HELPERS
  // ==========================================================================
  var STATUS_CONFIG = {
    'Register': { className: 'status-register' },
    'In Progress': { className: 'status-inprogress' },
    'Review': { className: 'status-review' },
    'Done': { className: 'status-done' }
  };

  function sanitize(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function highlightMatches(text, query) {
    var safeText = sanitize(text);
    if (!query || !query.trim()) return safeText;
    var safeQuery = escapeRegex(query.trim());
    var regex = new RegExp('(' + safeQuery + ')', 'gi');
    return safeText.replace(regex, '<mark class="highlight-text">$1</mark>');
  }

  function isKhmer(str) {
    return /[\u1780-\u17FF\u19E0-\u19FF]/.test(str || '');
  }

  function getLangAttr(str) {
    return isKhmer(str) ? ' lang="km"' : '';
  }

  function isLastStatus(statusName) {
    if (!statusName) return false;
    if (statusName.toLowerCase() === 'done') return true;
    if (DataService.statuses && DataService.statuses.length > 0) {
      return statusName === DataService.statuses[DataService.statuses.length - 1];
    }
    return false;
  }

  function isOverdue(deadlineStr, status) {
    if (!deadlineStr || status === 'Done') return false;
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var deadline = new Date(deadlineStr.includes('T') ? deadlineStr : deadlineStr + 'T00:00:00');
    return deadline < today;
  }

  function formatDate(isoOrDateStr) {
    if (!isoOrDateStr) return '-';
    try {
      var str = String(isoOrDateStr).trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        str = str + 'T00:00:00';
      }
      var d = new Date(str);
      if (isNaN(d.getTime())) return sanitize(isoOrDateStr);
      return d.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
      });
    } catch (e) {
      return sanitize(isoOrDateStr);
    }
  }

  function formatDateTime(isoStr) {
    if (!isoStr) return '-';
    try {
      var str = String(isoStr).trim();
      if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/.test(str)) {
        str = str.replace(' ', 'T');
      }
      var d = new Date(str);
      if (isNaN(d.getTime())) return sanitize(isoStr);
      return d.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch (e) {
      return sanitize(isoStr);
    }
  }

  function getDaysAgo(isoOrDateStr) {
    if (!isoOrDateStr) return '';
    try {
      var str = String(isoOrDateStr).trim();
      if (!str) return '';
      if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/.test(str)) {
        str = str.replace(' ', 'T');
      } else if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        str = str + 'T00:00:00';
      }
      var created = new Date(str);
      if (isNaN(created.getTime())) {
        created = new Date(isoOrDateStr);
        if (isNaN(created.getTime())) return '';
      }
      var now = new Date();
      var nowDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      var createdDay = new Date(created.getFullYear(), created.getMonth(), created.getDate());
      var diffDays = Math.round((nowDay.getTime() - createdDay.getTime()) / (1000 * 60 * 60 * 24));
      if (diffDays < 0) diffDays = 0;
      if (diffDays === 1) {
        return '1 day ago';
      }
      return diffDays + ' days ago';
    } catch (e) {
      return '';
    }
  }

  function getInitials(name) {
    if (!name) return '--';
    var parts = name.trim().split(/\s+/);
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function wait(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  // ==========================================================================
  // 4. APPLICATION STATE & DOM ELEMENTS
  // ==========================================================================
  var state = {
    currentUser: null,       // { name, pin, team } in RAM only - NEVER stored in localStorage
    selectedFilter: 'ALL',   // 'ALL' | memberName | 'COMPLETED' | 'ARCHIVED'
    filterType: 'ALL',       // 'ALL' | 'PERSON' | 'COMPLETED' | 'ARCHIVED'
    selectedTaskId: null,
    searchQuery: '',
    onlyMeFilter: false,
    syncTimer: null,              // 60-second pending queue check
    sessionCheckTimer: null,      // 10-minute session validity check
    consecutiveSyncAuthFailures: 0,
    startupInProgress: false,
    undoTimeout: null,
    lastUndoAction: null,    // { taskId, prevState, description }
    contextMenuTaskId: null,
    isCreatingTask: false,
    isUpdatingStatus: false,
    isDragging: false,
    pendingStateTasks: new Set(),
    lastRenderedTaskId: null,
    modalAssigneePrevious: null,
    editingTaskId: null
  };

  var el = {
    // Login Screen
    loginScreen: document.getElementById('login-screen'),
    loginForm: document.getElementById('login-form'),
    loginName: document.getElementById('username'),
    loginPin: document.getElementById('password'),
    loginNameError: document.getElementById('login-name-error'),
    loginPinError: document.getElementById('login-pin-error'),
    btnLogin: document.getElementById('btn-login'),

    // Startup Screen
    startupScreen: document.getElementById('startup-screen'),
    startupStatus: document.getElementById('startup-status'),
    startupErrorBox: document.getElementById('startup-error-box'),
    startupErrorMsg: document.getElementById('startup-error-msg'),
    btnStartupRetry: document.getElementById('btn-startup-retry'),
    btnStartupLogout: document.getElementById('btn-startup-logout'),
    stepNodes: [
      document.getElementById('step-node-1'),
      document.getElementById('step-node-2'),
      document.getElementById('step-node-3')
    ],
    connectors: [
      document.getElementById('connector-1'),
      document.getElementById('connector-2')
    ],
    travelDots: [
      document.getElementById('travel-dot-1'),
      document.getElementById('travel-dot-2')
    ],

    // Main App & Panels
    appContainer: document.getElementById('app-container'),
    panelPeople: document.getElementById('panel-people'),
    panelTasks: document.getElementById('panel-tasks'),
    panelDetails: document.getElementById('panel-details'),

    // Header & User Menu
    currentUserName: document.getElementById('current-user-name'),
    currentUserAvatar: document.getElementById('current-user-avatar'),
    btnLogout: document.getElementById('btn-logout'),

    // Left Navigation
    filterAll: document.getElementById('filter-all'),
    countAll: document.getElementById('count-all'),
    filterPersonal: document.getElementById('filter-personal'),
    countPersonal: document.getElementById('count-personal'),
    membersNavList: document.getElementById('members-nav-list'),
    membersEmptyState: document.getElementById('members-empty-state'),
    filterCompleted: document.getElementById('filter-completed'),
    countCompleted: document.getElementById('count-completed'),
    filterArchived: document.getElementById('filter-archived'),
    countArchived: document.getElementById('count-archived'),

    // Middle Tasks Panel
    currentViewTitle: document.getElementById('current-view-title'),
    viewTaskCount: document.getElementById('view-task-count'),
    syncIndicator: document.getElementById('sync-indicator'),
    syncDot: document.getElementById('sync-dot'),
    syncStatusText: document.getElementById('sync-status-text'),
    btnOpenNewTask: document.getElementById('btn-open-new-task'),
    taskSearchInput: document.getElementById('task-search-input'),
    btnClearSearch: document.getElementById('btn-clear-search'),
    searchSummaryBar: document.getElementById('search-summary-bar'),
    personalFilterBar: document.getElementById('personal-filter-bar'),
    filterOnlyMeCheckbox: document.getElementById('filter-only-me-checkbox'),
    tasksScrollContainer: document.getElementById('tasks-scroll-container'),
    taskItemsList: document.getElementById('task-items-list'),
    tasksEmptyState: document.getElementById('tasks-empty-state'),
    tasksEmptyTitle: document.getElementById('tasks-empty-title'),
    tasksEmptyDesc: document.getElementById('tasks-empty-desc'),

    // Right Details Panel
    detailsEmptyState: document.getElementById('details-empty-state'),
    detailsContent: document.getElementById('details-content'),
    detailTaskCode: document.getElementById('detail-task-code'),
    detailTaskName: document.getElementById('detail-task-name'),
    detailStatusBadge: document.getElementById('detail-status-badge'),
    detailPrivacyBadge: document.getElementById('detail-privacy-badge'),
    detailPrivacyToggleWrapper: document.getElementById('detail-privacy-toggle-wrapper'),
    detailShareTeamCheckbox: document.getElementById('detail-share-team-checkbox'),
    detailOverduePill: document.getElementById('detail-overdue-pill'),
    btnMarkComplete: document.getElementById('btn-mark-complete'),
    btnRestoreTask: document.getElementById('btn-restore-task'),
    btnDeleteTask: document.getElementById('btn-delete-task'),
    detailAssigneeName: document.getElementById('detail-assignee-name'),
    detailAssigneeAvatar: document.getElementById('detail-assignee-avatar'),
    detailDeadlineText: document.getElementById('detail-deadline-text'),
    detailCreatedBy: document.getElementById('detail-created-by'),
    detailCreatedAt: document.getElementById('detail-created-at'),
    detailDescription: document.getElementById('detail-description'),
    statusUpdateForm: document.getElementById('status-update-form'),
    updateStatusWrapper: document.getElementById('update-status-wrapper'),
    updateStatusSelect: document.getElementById('update-status-select'),
    updateStatusTrigger: document.getElementById('update-status-trigger'),
    updateStatusLabel: document.getElementById('update-status-label'),
    updateStatusMenu: document.getElementById('update-status-menu'),
    updateNoteInput: document.getElementById('update-note-input'),
    btnSubmitStatusUpdate: document.getElementById('btn-submit-status-update'),
    detailTimelineList: document.getElementById('detail-timeline-list'),
    historyEmptyState: document.getElementById('history-empty-state'),

    // Modal
    newTaskModal: document.getElementById('new-task-modal'),
    newTaskForm: document.getElementById('new-task-form'),
    modalHeading: document.getElementById('modal-heading'),
    taskNameInput: document.getElementById('task-name-input'),
    taskAssigneeInput: document.getElementById('task-assignee-input'),
    taskDeadlineInput: document.getElementById('task-deadline-input'),
    taskDeadlineHint: document.getElementById('task-deadline-hint'),
    taskEditStatusGroup: document.getElementById('task-edit-status-group'),
    taskStatusInput: document.getElementById('task-status-input'),
    taskDescInput: document.getElementById('task-desc-input'),
    taskShareTeamInput: document.getElementById('task-share-team-input'),
    taskModalInfoCallout: document.getElementById('task-modal-info-callout'),
    taskNameError: document.getElementById('task-name-error'),
    taskAssigneeError: document.getElementById('task-assignee-error'),
    taskDeadlineError: document.getElementById('task-deadline-error'),
    taskDescError: document.getElementById('task-desc-error'),
    btnSaveTask: document.getElementById('btn-save-task'),
    btnCloseModal: document.getElementById('btn-close-modal'),
    btnCancelModal: document.getElementById('btn-cancel-modal'),

    // Context Menu & Toast
    taskContextMenu: document.getElementById('task-context-menu'),
    ctxOpenTask: document.getElementById('ctx-open-task'),
    ctxEditTask: document.getElementById('ctx-edit-task'),
    ctxChangeStatus: document.getElementById('ctx-change-status'),
    ctxMarkComplete: document.getElementById('ctx-mark-complete'),
    ctxRestoreTask: document.getElementById('ctx-restore-task'),
    ctxDeleteTask: document.getElementById('ctx-delete-task'),
    undoToast: document.getElementById('undo-toast'),
    undoToastMessage: document.getElementById('undo-toast-message'),
    btnUndoToast: document.getElementById('btn-undo-toast'),

    // Confirm / Alert Modal
    confirmModal: document.getElementById('confirm-modal'),
    confirmModalTitle: document.getElementById('confirm-modal-title'),
    confirmModalHeadline: document.getElementById('confirm-modal-headline'),
    confirmModalMessage: document.getElementById('confirm-modal-message'),
    confirmModalIconWrap: document.getElementById('confirm-modal-icon-wrap'),
    btnCloseConfirmModal: document.getElementById('btn-close-confirm-modal'),
    btnCancelConfirm: document.getElementById('btn-cancel-confirm'),
    btnOkConfirm: document.getElementById('btn-ok-confirm'),

    // Version & Disclaimer
    appVersionIndicator: document.getElementById('app-version-indicator'),
    btnOpenDisclaimer: document.getElementById('btn-open-disclaimer'),
    disclaimerModal: document.getElementById('disclaimer-modal'),
    btnCloseDisclaimerModal: document.getElementById('btn-close-disclaimer-modal'),
    btnCloseDisclaimerBottom: document.getElementById('btn-close-disclaimer-bottom')
  };

  // ==========================================================================
  // 5. STEPPER ANIMATION & MANDATORY STARTUP FLOW
  // ==========================================================================
  function resetStepperUI() {
    el.stepNodes.forEach(function (node) {
      if (node) node.className = 'stepper-step';
    });
    el.connectors.forEach(function (c) {
      if (c) c.className = 'step-connector';
    });
    el.travelDots.forEach(function (dot) {
      if (dot) dot.className = 'travel-dot';
    });
    el.startupErrorBox.classList.add('hidden');
    el.startupStatus.textContent = 'Just a minute…';
  }

  function setStepActive(stepNum) {
    for (var i = 1; i <= 3; i++) {
      var node = el.stepNodes[i - 1];
      if (!node) continue;
      if (i < stepNum) {
        node.className = 'stepper-step is-complete';
      } else if (i === stepNum) {
        node.className = 'stepper-step is-active';
      } else {
        node.className = 'stepper-step';
      }
    }

    el.connectors.forEach(function (c, idx) {
      if (!c) return;
      if (idx + 1 < stepNum) {
        c.className = 'step-connector is-complete';
      } else if (idx + 1 === stepNum) {
        c.className = 'step-connector is-active';
      } else {
        c.className = 'step-connector';
      }
    });
  }

  function startGlidingDot(connIndex, speedMode) {
    var dot = el.travelDots[connIndex];
    if (!dot) return;
    dot.className = 'travel-dot';
    void dot.offsetWidth;
    if (speedMode === 'fast') {
      dot.classList.add('glide-fast');
    } else if (speedMode === 'loop') {
      dot.classList.add('glide-loop');
    } else {
      dot.classList.add('glide-normal');
    }
  }

  function stopGlidingDot(connIndex) {
    var dot = el.travelDots[connIndex];
    if (dot) {
      dot.className = 'travel-dot';
    }
  }

  async function startStartupFlow(preloadedData) {
    if (state.startupInProgress) return;
    state.startupInProgress = true;

    // Show loading screen; ensure main app and login screen are hidden
    el.loginScreen.classList.add('hidden');
    el.appContainer.classList.add('hidden');
    el.startupScreen.classList.remove('hidden', 'fade-out');
    resetStepperUI();

    var currentStep = 1;
    var dataPayload = preloadedData || null;
    var dataFailed = false;
    var dataErrorMessage = '';
    var dataCompleted = !!preloadedData;

    // Data Promise with 20-second timeout
    var dataPromise = (async function () {
      if (preloadedData) {
        return preloadedData;
      }
      try {
        var timeoutPromise = new Promise(function (_, reject) {
          setTimeout(function () {
            reject(new Error('Connection timed out after 20 seconds. Backend took too long to respond.'));
          }, 20000);
        });

        var startPromise = apiCall('start', { forceRefresh: true });
        var res = await Promise.race([startPromise, timeoutPromise]);

        if (res && (res.error === 'unauthorized' || res.error === 'locked' || (res.ok === false && (res.error === 'unauthorized' || res.error === 'locked')))) {
          logout();
          var authErr = (res.error === 'locked')
            ? 'Account is temporarily locked due to 5 failed attempts (15 minutes).'
            : 'Session unauthorized or credentials invalid. Please sign in again.';
          dataFailed = true;
          dataErrorMessage = authErr;
          dataCompleted = true;
          throw new Error(authErr);
        }

        if (res && res.success && res.data) {
          dataPayload = res.data;
          dataCompleted = true;
          return res.data;
        } else {
          var errMsg = (res && res.message) || 'Unable to download team data from Google Sheets.';
          dataFailed = true;
          dataErrorMessage = errMsg;
          dataCompleted = true;
          throw new Error(errMsg);
        }
      } catch (err) {
        dataFailed = true;
        dataErrorMessage = err.message || 'Unable to connect to Google Sheets backend.';
        dataCompleted = true;
        throw err;
      }
    })();

    // Coordinated Stepper Animation
    try {
      // Step 1: Your Device
      currentStep = 1;
      setStepActive(1);
      el.startupStatus.textContent = 'Just a minute…';

      startGlidingDot(0, dataCompleted ? 'fast' : 'normal');
      await wait(dataCompleted ? 350 : 700);

      if (!dataCompleted) {
        startGlidingDot(0, 'loop');
        while (!dataCompleted) {
          await wait(200);
        }
      }
      stopGlidingDot(0);

      // Check if data fetch failed at Step 1/2
      if (dataFailed) {
        throw new Error(dataErrorMessage);
      }

      // Step 2: Cloud Data
      currentStep = 2;
      setStepActive(2);
      el.startupStatus.textContent = 'Downloading team data…';

      startGlidingDot(1, dataCompleted ? 'fast' : 'normal');
      await wait(dataCompleted ? 350 : 700);

      if (!dataCompleted) {
        startGlidingDot(1, 'loop');
        while (!dataCompleted) {
          await wait(200);
        }
      }
      stopGlidingDot(1);

      if (dataFailed) {
        throw new Error(dataErrorMessage);
      }

      // Step 3: Dashboard
      currentStep = 3;
      setStepActive(3);
      el.startupStatus.textContent = 'Building your dashboard…';
      await wait(400);

      // Await data payload resolution
      await dataPromise;

      // Ingest data into DataService
      DataService.loadSyncData(dataPayload);

      // Smooth fade into the main app
      el.startupStatus.textContent = 'Ready!';
      await wait(250);

      el.startupScreen.classList.add('fade-out');
      await wait(280);

      el.startupScreen.classList.add('hidden');
      el.startupScreen.classList.remove('fade-out');
      el.appContainer.classList.remove('hidden');

      state.startupInProgress = false;

      // Render Dashboard
      renderAll();
      startPeriodicSync();
      startSessionValidityCheck();
      flushPendingActions();

    } catch (failureErr) {
      state.startupInProgress = false;
      stopGlidingDot(0);
      stopGlidingDot(1);

      // Mark the step where failure occurred
      var failedStepNode = el.stepNodes[Math.min(currentStep, 3) - 1];
      if (failedStepNode) failedStepNode.className = 'stepper-step is-error';

      el.startupStatus.textContent = 'Connection failed';
      el.startupErrorMsg.textContent = failureErr.message || 'Unable to connect to Google Sheets backend.';
      el.startupErrorBox.classList.remove('hidden');
    }
  }

  // ==========================================================================
  // 6. BACKGROUND QUIET SYNC (60s QUEUE CHECK & 10m SESSION CHECK)
  // ==========================================================================
  function updateSyncStatus(status) {
    if (!el.syncIndicator) return;
    el.syncIndicator.className = 'sync-indicator ' + status;
    if (status === 'saving') {
      el.syncStatusText.textContent = 'Saving…';
    } else if (status === 'syncing') {
      el.syncStatusText.textContent = 'Syncing…';
    } else if (status === 'error') {
      el.syncStatusText.textContent = "Couldn't Sync";
    } else {
      el.syncStatusText.textContent = 'Synced';
    }
  }

  /**
   * 60-Second Background Sync:
   * 1. Checks local pending-actions queue first.
   * 2. If queue is empty: DO NOTHING. Zero requests sent!
   * 3. If queue has items AND user is not mid-edit: syncs the queue (send all pending actions together).
   * 4. If user is mid-edit at that moment: skips this cycle and checks again on next 60s tick.
   */
  async function syncQuietly() {
    if (!state.currentUser || state.startupInProgress) return;

    // 1. Check local pending-actions queue first
    var pending = getPendingActions();
    if (!pending || pending.length === 0) {
      // Empty queue: DO NOTHING. No request sent, no server call at all.
      return;
    }

    // 2. If user is actively mid-edit, skip this cycle and check on next 60s tick
    if (isUserMidEdit()) {
      return;
    }

    // 3. Sync the queue now — send all pending actions together
    await syncPendingQueue();
  }

  function startPeriodicSync() {
    if (state.syncTimer) clearInterval(state.syncTimer);
    state.syncTimer = setInterval(function () {
      syncQuietly();
    }, 60000);
  }

  /**
   * Separate, less frequent Session-Validity Check (every 10 minutes):
   * Runs regardless of queue state since it's a security check, not a data sync.
   * Does not merge into the 60s queue check.
   */
  async function checkSessionValidity() {
    if (!state.currentUser || state.startupInProgress) return;

    var res = await apiCall('sync', { lastSync: DataService.lastSync, checkOnly: true });

    // 1. Account locked: prompt re-login / locked modal immediately
    if (res && res.error === 'locked') {
      state.consecutiveSyncAuthFailures = 0;
      handleAuthFailureResponse(res);
      return;
    }

    // 2. Transient errors (server_busy, network_error, timeout): NEVER log out
    if (!res || res.error === 'server_busy' || res.error === 'network_error' || res.error === 'busy') {
      return;
    }

    // 3. Genuine unauthorized: require 3 consecutive occurrences before kicking out
    if (res.error === 'unauthorized' || (res.ok === false && res.error === 'unauthorized')) {
      state.consecutiveSyncAuthFailures = (state.consecutiveSyncAuthFailures || 0) + 1;
      if (state.consecutiveSyncAuthFailures >= 3) {
        state.consecutiveSyncAuthFailures = 0;
        handleAuthFailureResponse(res);
      }
      return;
    }

    // 4. Success: reset failure counter, and load new data if delta available
    state.consecutiveSyncAuthFailures = 0;
    if (res && res.success && res.data) {
      DataService.loadSyncData(res.data);
      renderAll(true);
    }
  }

  function startSessionValidityCheck() {
    if (state.sessionCheckTimer) clearInterval(state.sessionCheckTimer);
    state.sessionCheckTimer = setInterval(function () {
      checkSessionValidity();
    }, 10 * 60 * 1000); // 10 minutes (less frequent security check)
  }

  // ==========================================================================
  // 7. AUTHENTICATION & LOGIN FLOW (Password Manager Friendly)
  // ==========================================================================
  function initAuth() {
    if (!API_URL) {
      console.error('Missing API configuration: CONFIG_API_URL is missing or unconfigured.');
      if (el.loginPinError) {
        el.loginPinError.textContent = 'Missing API configuration. Please set CONFIG_API_URL in config.js.';
      }
    } else {
      // Ping backend immediately on page load to wake up script
      apiCall('ping');
    }

    // Populate app version from config.js
    if (el.appVersionIndicator) {
      el.appVersionIndicator.textContent = (typeof APP_VERSION !== 'undefined' && APP_VERSION) ? APP_VERSION : 'v1.0.0';
    }

    // Never show cached data before login is verified by Code.gs.
    // Credentials kept in memory or sessionStorage only, NEVER in localStorage.
    var savedSession = null;
    try {
      var raw = sessionStorage.getItem('team_tasks_session_user');
      if (raw) savedSession = JSON.parse(raw);
    } catch (e) {}

    // Clean any legacy credentials from localStorage
    try {
      localStorage.removeItem('team_tasks_session_user');
    } catch (e) {}

    if (savedSession && savedSession.name && savedSession.pin) {
      state.currentUser = savedSession;
      el.currentUserName.textContent = state.currentUser.name;
      if (el.currentUserAvatar) el.currentUserAvatar.textContent = getInitials(state.currentUser.name);
      el.loginScreen.classList.add('hidden');
      startStartupFlow();
    } else {
      showLogin();
    }
  }

  function showLogin() {
    if (state.syncTimer) clearInterval(state.syncTimer);
    if (state.sessionCheckTimer) clearInterval(state.sessionCheckTimer);
    el.loginScreen.classList.remove('hidden');
    el.startupScreen.classList.add('hidden');
    el.appContainer.classList.add('hidden');
    el.loginName.value = '';
    el.loginPin.value = '';
    el.loginNameError.textContent = '';
    el.loginPinError.textContent = '';
    el.btnLogin.disabled = false;
    el.btnLogin.textContent = 'Sign In';
  }

  async function handleLoginSubmit(e) {
    e.preventDefault();
    el.loginNameError.textContent = '';
    el.loginPinError.textContent = '';

    var name = el.loginName.value.trim();
    var pin = el.loginPin.value.trim();
    var hasError = false;

    if (!name) {
      el.loginNameError.textContent = 'Please enter your name';
      hasError = true;
    }

    if (!pin) {
      el.loginPinError.textContent = 'Please enter your PIN code';
      hasError = true;
    }

    if (hasError) return;

    el.btnLogin.disabled = true;
    el.btnLogin.textContent = 'Signing In...';

    // Verify directly against Code.gs (sole source of trust)
    var res = await apiCall('start', { name: name, pin: pin, forceRefresh: true });

    el.btnLogin.disabled = false;
    el.btnLogin.textContent = 'Sign In';

    if (res && res.ok && res.success && res.user) {
      var sessionUser = {
        name: res.user.name,
        pin: pin,
        team: res.user.team
      };

      // Store credentials in memory and sessionStorage only - NEVER localStorage
      state.currentUser = sessionUser;
      try {
        sessionStorage.setItem('team_tasks_session_user', JSON.stringify(sessionUser));
      } catch (sessErr) {}
      try {
        localStorage.removeItem('team_tasks_session_user');
      } catch (e) {}

      el.currentUserName.textContent = state.currentUser.name;
      if (el.currentUserAvatar) el.currentUserAvatar.textContent = getInitials(state.currentUser.name);

      el.loginScreen.classList.add('hidden');

      // Mandatory startup flow with verified server data
      startStartupFlow(res.data);
    } else {
      el.loginPin.value = '';
      if (res && res.error === 'locked') {
        el.loginPinError.textContent = 'Account temporarily locked (5 failed attempts). Try again in 15 minutes.';
      } else if (res && res.error === 'missing_config') {
        el.loginPinError.textContent = res.message || 'Missing API configuration. Please configure CONFIG_API_URL in config.js.';
      } else {
        el.loginPinError.textContent = 'Invalid name or PIN. Please verify credentials.';
      }
      el.loginPin.focus();
    }
  }

  function logout() {
    if (state.syncTimer) clearInterval(state.syncTimer);
    if (state.sessionCheckTimer) clearInterval(state.sessionCheckTimer);
    state.currentUser = null;
    state.selectedTaskId = null;
    state.lastRenderedTaskId = null;
    state.onlyMeFilter = false;
    if (el.filterOnlyMeCheckbox) el.filterOnlyMeCheckbox.checked = false;
    DataService.tasks = [];
    DataService.history = [];
    DataService.statuses = ['Register', 'In Progress', 'Review', 'Done'];
    DataService.users = [];
    DataService.lastSync = '';
    try {
      sessionStorage.clear();
    } catch (e) {}
    try {
      localStorage.removeItem('team_tasks_session_user');
      localStorage.removeItem('team_tasks_local_data');
    } catch (e) {}
    showLogin();
  }

  // ==========================================================================
  // 8. MAIN APP RENDERING
  // ==========================================================================
  function renderAll(preserveDrafts) {
    renderPeoplePanel();
    renderTasksList();
    var isTaskSwitch = (preserveDrafts === false);
    renderTaskDetails(isTaskSwitch);
  }

  function renderPeoplePanel() {
    var allTasks = DataService.getTasks();
    var publicTasks = allTasks.filter(function (t) { return !t.isPrivate; });
    var activeTasks = publicTasks.filter(function (t) { return (t.state || 'Active') === 'Active'; });
    var completedTasks = publicTasks.filter(function (t) { return t.state === 'Completed'; });
    var archivedTasks = publicTasks.filter(function (t) { return t.state === 'Archived'; });

    var currentUserName = (state.currentUser ? state.currentUser.name : '').trim().toLowerCase();
    var personalTasks = allTasks.filter(function (t) {
      return (t.state || 'Active') === 'Active' &&
             (t.createdBy || '').trim().toLowerCase() === currentUserName;
    });

    el.countAll.textContent = activeTasks.length;
    if (el.countPersonal) el.countPersonal.textContent = personalTasks.length;
    if (el.countCompleted) el.countCompleted.textContent = completedTasks.length;
    if (el.countArchived) el.countArchived.textContent = archivedTasks.length;

    // Active state highlighting on left panel items
    el.filterAll.classList.toggle('active', state.filterType === 'ALL');
    if (el.filterPersonal) el.filterPersonal.classList.toggle('active', state.filterType === 'PERSONAL');
    el.filterCompleted.classList.toggle('active', state.filterType === 'COMPLETED');
    el.filterArchived.classList.toggle('active', state.filterType === 'ARCHIVED');

    el.membersNavList.innerHTML = '';

    if (DataService.users.length === 0) {
      el.membersEmptyState.textContent = 'No team members loaded';
      el.membersEmptyState.classList.remove('hidden');
      return;
    }

    var renderedMembersCount = 0;

    DataService.users.forEach(function (user) {
      var userCount = activeTasks.filter(function (t) {
        return (t.assignedTo || '').toLowerCase() === user.name.toLowerCase();
      }).length;

      if (userCount === 0) {
        return;
      }
      renderedMembersCount++;

      var isActive = state.filterType === 'PERSON' && state.selectedFilter.toLowerCase() === user.name.toLowerCase();
      var nameLang = getLangAttr(user.name);

      var itemBtn = document.createElement('button');
      itemBtn.type = 'button';
      itemBtn.className = 'person-nav-item' + (isActive ? ' active' : '');
      itemBtn.innerHTML = `
        <div class="person-nav-left">
          <span class="person-name"${nameLang}>${sanitize(user.name)}</span>
        </div>
        <span class="count-badge">${userCount}</span>
      `;

      itemBtn.addEventListener('click', function () {
        setPersonFilter(user.name);
      });

      el.membersNavList.appendChild(itemBtn);
    });

    if (renderedMembersCount === 0) {
      el.membersEmptyState.textContent = 'No team members with active tasks';
      el.membersEmptyState.classList.remove('hidden');
    } else {
      el.membersEmptyState.classList.add('hidden');
    }
  }

  function getFilteredTasks() {
    var tasks = DataService.getTasks();

    // 1. Filter by State / View
    if (state.filterType === 'COMPLETED') {
      tasks = tasks.filter(function (t) { return t.state === 'Completed'; });
    } else if (state.filterType === 'ARCHIVED') {
      tasks = tasks.filter(function (t) { return t.state === 'Archived'; });
    } else if (state.filterType === 'PERSONAL') {
      var currentUserName = (state.currentUser ? state.currentUser.name : '').trim().toLowerCase();
      tasks = tasks.filter(function (t) {
        var isMine = (t.state || 'Active') === 'Active' &&
               (t.createdBy || '').trim().toLowerCase() === currentUserName;
        if (!isMine) return false;
        if (state.onlyMeFilter) {
          return t.isPrivate === true;
        }
        return true;
      });
    } else if (state.filterType === 'PERSON') {
      tasks = tasks.filter(function (t) {
        return (t.state || 'Active') === 'Active' &&
               (t.assignedTo || '').toLowerCase() === state.selectedFilter.toLowerCase();
      });
    } else {
      // 'ALL' default: active tasks
      tasks = tasks.filter(function (t) { return (t.state || 'Active') === 'Active'; });
    }

    // 2. Universal Search Filter (all fields searched at once)
    var q = state.searchQuery.trim().toLowerCase();
    if (q) {
      tasks = tasks.filter(function (t) {
        var code = (t.taskCode || generateTaskCodeFromDate(t.createdAt, t.taskId)).toLowerCase();
        var codeMatch = code.includes(q);
        var nameMatch = (t.taskName || '').toLowerCase().includes(q);
        var descMatch = (t.description || '').toLowerCase().includes(q);
        var assigneeMatch = (t.assignedTo || '').toLowerCase().includes(q);
        var creatorMatch = (t.createdBy || '').toLowerCase().includes(q);
        var statusMatch = (t.status || '').toLowerCase().includes(q);
        var deadlineMatch = (t.deadline || '').toLowerCase().includes(q);

        // Also search in history notes for this task
        var taskHistory = DataService.getHistoryByTaskId(t.taskId);
        var historyMatch = taskHistory.some(function (h) {
          return (h.note || '').toLowerCase().includes(q) ||
                 (h.updatedBy || '').toLowerCase().includes(q);
        });

        return codeMatch || nameMatch || descMatch || assigneeMatch || creatorMatch || statusMatch || deadlineMatch || historyMatch;
      });
    }

    return tasks;
  }

  function getSearchMatchTag(task, q) {
    if (!q) return '';
    q = q.toLowerCase();
    var taskCode = (task.taskCode || generateTaskCodeFromDate(task.createdAt, task.taskId)).toLowerCase();
    if (taskCode.includes(q)) return 'Code';
    if ((task.taskName || '').toLowerCase().includes(q)) return 'Title';
    if ((task.assignedTo || '').toLowerCase().includes(q)) return 'Assignee';
    if ((task.description || '').toLowerCase().includes(q)) return 'Description';
    if ((task.status || '').toLowerCase().includes(q)) return 'Status';
    if ((task.createdBy || '').toLowerCase().includes(q)) return 'Creator';

    var taskHistory = DataService.getHistoryByTaskId(task.taskId);
    if (taskHistory.some(function (h) { return (h.note || '').toLowerCase().includes(q); })) {
      return 'Note';
    }
    return 'Match';
  }

  function renderTasksList() {
    var filteredTasks = getFilteredTasks();
    var allTasks = DataService.getTasks();

    // Set View Title
    if (state.filterType === 'COMPLETED') {
      el.currentViewTitle.textContent = 'Completed Tasks';
    } else if (state.filterType === 'ARCHIVED') {
      el.currentViewTitle.textContent = 'Deleted Tasks';
    } else if (state.filterType === 'PERSONAL') {
      el.currentViewTitle.textContent = 'Personal Tasks';
    } else if (state.filterType === 'PERSON') {
      el.currentViewTitle.textContent = state.selectedFilter + "'s Tasks";
    } else {
      el.currentViewTitle.textContent = 'All Tasks';
    }

    el.viewTaskCount.textContent = filteredTasks.length + (filteredTasks.length === 1 ? ' task' : ' tasks');

    // Personal Tasks Only Me filter bar
    if (el.personalFilterBar) {
      if (state.filterType === 'PERSONAL') {
        el.personalFilterBar.classList.remove('hidden');
        if (el.filterOnlyMeCheckbox) {
          el.filterOnlyMeCheckbox.checked = !!state.onlyMeFilter;
        }
      } else {
        el.personalFilterBar.classList.add('hidden');
      }
    }

    // Search summary bar
    if (state.searchQuery.trim().length > 0) {
      el.searchSummaryBar.classList.remove('hidden');
      el.searchSummaryBar.innerHTML = `
        <span>Searching all tasks: <strong>${filteredTasks.length}</strong> found for "${sanitize(state.searchQuery)}"</span>
        <button type="button" class="btn btn-ghost btn-sm" id="btn-clear-search-summary">Clear</button>
      `;
      var btnClearSum = document.getElementById('btn-clear-search-summary');
      if (btnClearSum) {
        btnClearSum.addEventListener('click', function () {
          state.searchQuery = '';
          el.taskSearchInput.value = '';
          el.btnClearSearch.classList.add('hidden');
          el.searchSummaryBar.classList.add('hidden');
          renderTasksList();
        });
      }
    } else {
      el.searchSummaryBar.classList.add('hidden');
    }

    // Empty State
    if (filteredTasks.length === 0) {
      el.taskItemsList.innerHTML = '';
      el.tasksEmptyState.classList.remove('hidden');

      if (state.searchQuery.trim().length > 0) {
        el.tasksEmptyTitle.textContent = 'No matching tasks';
        el.tasksEmptyDesc.textContent = 'No tasks match "' + sanitize(state.searchQuery) + '". Press Esc to clear.';
      } else if (state.filterType === 'COMPLETED') {
        el.tasksEmptyTitle.textContent = 'No completed tasks';
        el.tasksEmptyDesc.textContent = 'Finished tasks marked as Complete will appear here.';
      } else if (state.filterType === 'ARCHIVED') {
        el.tasksEmptyTitle.textContent = 'No deleted tasks';
        el.tasksEmptyDesc.textContent = 'Tasks moved to Deleted will appear here.';
      } else if (state.filterType === 'PERSONAL') {
        if (state.onlyMeFilter) {
          el.tasksEmptyTitle.textContent = 'No Only Me tasks';
          el.tasksEmptyDesc.textContent = 'You have no private tasks visible only to you.';
        } else {
          el.tasksEmptyTitle.textContent = 'No personal tasks yet';
          el.tasksEmptyDesc.textContent = 'Tasks you create will appear here.';
        }
      } else if (state.filterType === 'PERSON') {
        el.tasksEmptyTitle.textContent = 'No tasks for this member';
        el.tasksEmptyDesc.textContent = 'There are no active tasks assigned to ' + sanitize(state.selectedFilter) + '.';
      } else {
        el.tasksEmptyTitle.textContent = 'No tasks yet';
        el.tasksEmptyDesc.textContent = 'Click New Task to start.';
      }
      return;
    }

    el.tasksEmptyState.classList.add('hidden');

    // Build task rows with DocumentFragment for speed
    var fragment = document.createDocumentFragment();

    filteredTasks.forEach(function (task) {
      var isSelected = task.taskId === state.selectedTaskId;
      var overdue = isOverdue(task.deadline, task.status);
      var statusMeta = STATUS_CONFIG[task.status] || { className: 'status-register' };
      var lastClass = isLastStatus(task.status) ? ' status-last' : '';
      var nameLang = getLangAttr(task.taskName);
      var assigneeLang = getLangAttr(task.assignedTo);

      var matchTag = state.searchQuery.trim().length > 0
        ? `<span class="search-tag">${getSearchMatchTag(task, state.searchQuery)}</span>`
        : '';

      var displayName = state.searchQuery.trim().length > 0
        ? highlightMatches(task.taskName, state.searchQuery)
        : sanitize(task.taskName);

      var displayAssignee = state.searchQuery.trim().length > 0
        ? highlightMatches(task.assignedTo, state.searchQuery)
        : sanitize(task.assignedTo);

      var daysText = getDaysAgo(task.createdAt);
      var daysBadge = daysText
        ? `<span class="task-age-badge" lang="en" title="Created: ${sanitize(formatDateTime(task.createdAt))}">${daysText}</span>`
        : '';

      var row = document.createElement('div');
      row.className = 'task-item-row' + (isSelected ? ' selected' : '');
      row.dataset.taskId = task.taskId;
      row.tabIndex = 0;

      row.innerHTML = `
        <div class="task-row-top">
          <h3 class="task-row-title"${nameLang}>
            ${daysBadge}${displayName}
            ${matchTag}
          </h3>
          <div class="task-row-badges">
            ${task.isPrivate ? '<span class="privacy-badge" title="Visible only to you"><svg class="icon icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="0" ry="0"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg><span>Only Me</span></span>' : ''}
            <span class="status-badge ${statusMeta.className}${lastClass}">${sanitize(task.status)}</span>
          </div>
        </div>
        <div class="task-row-bottom">
          <span class="task-assignee-chip">
            <span${assigneeLang}>${displayAssignee}</span>
          </span>
          <div class="task-row-meta-right">
            <span class="task-deadline ${overdue ? 'is-overdue' : ''}">
              <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
                <line x1="16" y1="2" x2="16" y2="6"/>
                <line x1="8" y1="2" x2="8" y2="6"/>
                <line x1="3" y1="10" x2="21" y2="10"/>
              </svg>
              <span>${sanitize(formatDate(task.deadline))}${overdue ? ' (Overdue)' : ''}</span>
            </span>
          </div>
        </div>
      `;

      row.addEventListener('click', function () {
        selectTask(task.taskId);
      });

      row.addEventListener('contextmenu', function (e) {
        e.preventDefault();
        openContextMenu(e, task.taskId);
      });

      fragment.appendChild(row);
    });

    var prevScroll = el.taskItemsList ? el.taskItemsList.scrollTop : 0;
    el.taskItemsList.innerHTML = '';
    el.taskItemsList.appendChild(fragment);
    if (el.taskItemsList && prevScroll) {
      el.taskItemsList.scrollTop = prevScroll;
    }
  }

  function renderTaskDetails(isTaskSwitch) {
    if (!state.selectedTaskId) {
      state.lastRenderedTaskId = null;
      el.detailsEmptyState.classList.remove('hidden');
      el.detailsContent.classList.add('hidden');
      return;
    }

    var task = DataService.getTaskById(state.selectedTaskId);
    if (!task) {
      state.selectedTaskId = null;
      state.lastRenderedTaskId = null;
      el.detailsEmptyState.classList.remove('hidden');
      el.detailsContent.classList.add('hidden');
      return;
    }

    el.detailsEmptyState.classList.add('hidden');
    el.detailsContent.classList.remove('hidden');

    if (el.detailTaskCode) {
      var code = task.taskCode || generateTaskCodeFromDate(task.createdAt, task.taskId);
      el.detailTaskCode.textContent = code;
      el.detailTaskCode.setAttribute('title', 'Task Code: ' + code + ' (Click to copy)');
    }

    var daysText = getDaysAgo(task.createdAt);
    var daysHtml = daysText
      ? `<span class="task-age-badge" lang="en" title="Created: ${sanitize(formatDateTime(task.createdAt))}">${daysText}</span>`
      : '';
    el.detailTaskName.innerHTML = daysHtml + sanitize(task.taskName);
    if (isKhmer(task.taskName)) {
      el.detailTaskName.setAttribute('lang', 'km');
    } else {
      el.detailTaskName.removeAttribute('lang');
    }

    var statusMeta = STATUS_CONFIG[task.status] || { className: 'status-register' };
    var lastClass = isLastStatus(task.status) ? ' status-last' : '';
    el.detailStatusBadge.className = 'status-badge ' + statusMeta.className + lastClass;
    el.detailStatusBadge.textContent = task.status;

    if (el.detailPrivacyBadge) {
      if (task.isPrivate) {
        el.detailPrivacyBadge.classList.remove('hidden');
      } else {
        el.detailPrivacyBadge.classList.add('hidden');
      }
    }

    var currentUserName = (state.currentUser ? state.currentUser.name : '').trim().toLowerCase();
    var isCreator = task.createdBy && task.createdBy.trim().toLowerCase() === currentUserName;
    if (el.detailPrivacyToggleWrapper && el.detailShareTeamCheckbox) {
      if (isCreator) {
        el.detailPrivacyToggleWrapper.classList.remove('hidden');
        el.detailShareTeamCheckbox.checked = !task.isPrivate;
        if (!task.isPrivate) {
          el.detailPrivacyToggleWrapper.classList.add('is-shared');
          el.detailPrivacyToggleWrapper.classList.remove('is-private');
        } else {
          el.detailPrivacyToggleWrapper.classList.add('is-private');
          el.detailPrivacyToggleWrapper.classList.remove('is-shared');
        }
      } else {
        el.detailPrivacyToggleWrapper.classList.add('hidden');
      }
    }

    var overdue = isOverdue(task.deadline, task.status);
    if (overdue) {
      el.detailOverduePill.classList.remove('hidden');
    } else {
      el.detailOverduePill.classList.add('hidden');
    }

    // Toggle Complete / Restore / Delete button states depending on current task state
    if (task.state === 'Archived') {
      el.btnMarkComplete.classList.add('hidden');
      if (el.btnRestoreTask) el.btnRestoreTask.classList.remove('hidden');
      if (el.btnDeleteTask) {
        el.btnDeleteTask.classList.remove('hidden');
        el.btnDeleteTask.title = 'Permanently delete task';
      }
    } else if (task.state === 'Completed') {
      el.btnMarkComplete.classList.remove('hidden');
      el.btnMarkComplete.innerHTML = `
        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>
          <polyline points="15 3 21 3 21 9"/>
          <line x1="10" y1="14" x2="21" y2="3"/>
        </svg>
        <span>Active</span>
      `;
      el.btnMarkComplete.title = 'Reopen task as Active';

      if (el.btnRestoreTask) el.btnRestoreTask.classList.add('hidden');
      if (el.btnDeleteTask) {
        el.btnDeleteTask.classList.remove('hidden');
        el.btnDeleteTask.title = 'Move task to Deleted';
      }
    } else {
      el.btnMarkComplete.classList.remove('hidden');
      el.btnMarkComplete.innerHTML = `
        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="20 6 9 17 4 12"/>
        </svg>
        <span>Complete</span>
      `;
      el.btnMarkComplete.title = 'Mark task completed';

      if (el.btnRestoreTask) el.btnRestoreTask.classList.add('hidden');
      if (el.btnDeleteTask) {
        el.btnDeleteTask.classList.remove('hidden');
        el.btnDeleteTask.title = 'Move task to Deleted';
      }
    }

    el.detailAssigneeName.textContent = task.assignedTo || '-';
    if (isKhmer(task.assignedTo)) {
      el.detailAssigneeName.setAttribute('lang', 'km');
    } else {
      el.detailAssigneeName.removeAttribute('lang');
    }

    if (el.detailAssigneeAvatar) el.detailAssigneeAvatar.textContent = getInitials(task.assignedTo);
    el.detailDeadlineText.textContent = formatDate(task.deadline);
    if (overdue) {
      el.detailDeadlineText.textContent += ' (Overdue)';
      el.detailDeadlineText.style.color = 'var(--color-danger-text)';
      el.detailDeadlineText.style.fontWeight = '600';
    } else {
      el.detailDeadlineText.style.color = '';
      el.detailDeadlineText.style.fontWeight = '';
    }

    el.detailCreatedBy.textContent = task.createdBy || '-';
    if (isKhmer(task.createdBy)) {
      el.detailCreatedBy.setAttribute('lang', 'km');
    } else {
      el.detailCreatedBy.removeAttribute('lang');
    }

    el.detailCreatedAt.textContent = formatDateTime(task.createdAt);

    if (task.description && task.description.trim()) {
      el.detailDescription.textContent = task.description;
      if (isKhmer(task.description)) {
        el.detailDescription.setAttribute('lang', 'km');
      } else {
        el.detailDescription.removeAttribute('lang');
      }
    } else {
      el.detailDescription.innerHTML = '<em>No description provided.</em>';
      el.detailDescription.removeAttribute('lang');
    }

    // Form controls management:
    // Only treat as different task if explicitly switched (isTaskSwitch === true)
    // or if the task being rendered is different from state.lastRenderedTaskId
    var isDifferentTask = (isTaskSwitch === true) || (state.lastRenderedTaskId !== task.taskId);

    // Check if options in dropdown match current DataService.statuses
    var currentOptions = Array.from(el.updateStatusSelect.options).map(function (o) { return o.value; });
    var statusListMatches = (currentOptions.length === DataService.statuses.length) &&
      DataService.statuses.every(function (st, idx) { return st === currentOptions[idx]; });

    if (!statusListMatches || el.updateStatusSelect.options.length === 0) {
      var preservedStatus = isDifferentTask ? task.status : (el.updateStatusSelect.value || task.status);
      el.updateStatusSelect.innerHTML = '';
      DataService.statuses.forEach(function (st) {
        var opt = document.createElement('option');
        opt.value = st;
        opt.textContent = st;
        if (st === preservedStatus) opt.selected = true;
        el.updateStatusSelect.appendChild(opt);
      });
      el.updateStatusSelect.value = preservedStatus;
    } else if (isDifferentTask) {
      // Switching to a different task: set dropdown to the newly selected task's current status
      el.updateStatusSelect.value = task.status || (DataService.statuses[0] || 'Register');
    }
    // If NOT a different task (background sync or same-task re-render), NEVER touch el.updateStatusSelect.value!

    // Synchronize custom dropdown UI (matches website UI with uniform normal font weight)
    syncCustomStatusSelect();

    // Clear note input ONLY when explicitly switching to a different task
    if (isDifferentTask) {
      el.updateNoteInput.value = '';
    }
    // If NOT a different task (background sync or same-task re-render), NEVER touch el.updateNoteInput.value!

    state.lastRenderedTaskId = task.taskId;

    // Render Timeline with DocumentFragment
    var historyItems = DataService.getHistoryByTaskId(task.taskId);
    if (historyItems.length === 0 && task.createdAt) {
      historyItems = [{
        historyId: 'HIST-INIT-' + task.taskId,
        taskId: task.taskId,
        taskName: task.taskName,
        status: (DataService.statuses && DataService.statuses.length > 0) ? DataService.statuses[0] : (task.status || 'Register'),
        note: task.description ? ('Task created: ' + task.description) : 'Task created',
        updatedBy: task.createdBy || 'Team Member',
        timestamp: task.createdAt
      }];
    }
    el.detailTimelineList.innerHTML = '';

    if (historyItems.length === 0) {
      el.historyEmptyState.classList.remove('hidden');
    } else {
      el.historyEmptyState.classList.add('hidden');
      var timelineFrag = document.createDocumentFragment();

      historyItems.forEach(function (entry, index) {
        var entryStatus = entry.status || task.status || 'Register';
        var entryStatusMeta = STATUS_CONFIG[entryStatus] || { className: 'status-register' };
        var entryLastClass = isLastStatus(entryStatus) ? ' status-last' : '';
        var isFirstEntry = index === 0;

        // Highlight entry if search matched this history note
        var isNoteMatch = state.searchQuery.trim().length > 0 &&
                          (entry.note || '').toLowerCase().includes(state.searchQuery.trim().toLowerCase());

        var entryDiv = document.createElement('div');
        entryDiv.className = 'timeline-entry' + (isFirstEntry ? ' latest-entry' : '') + (isNoteMatch ? ' highlighted-entry' : '');

        var noteLang = getLangAttr(entry.note);
        var author = entry.updatedBy || entry.user || 'Team Member';
        var authorLang = getLangAttr(author);

        var displayNote = state.searchQuery.trim().length > 0
          ? highlightMatches(entry.note, state.searchQuery)
          : sanitize(entry.note);

        var noteHtml = entry.note && entry.note.trim()
          ? `<div class="timeline-note"${noteLang}>${displayNote}</div>`
          : '';

        entryDiv.innerHTML = `
          <div class="timeline-dot"></div>
          <div class="timeline-header">
            <span class="status-badge ${entryStatusMeta.className}${entryLastClass}">${sanitize(entryStatus)}</span>
            <span class="timeline-author"${authorLang}>${sanitize(author)}</span>
            <span class="timeline-time">${sanitize(formatDateTime(entry.timestamp))}</span>
          </div>
          ${noteHtml}
        `;
        timelineFrag.appendChild(entryDiv);
      });

      el.detailTimelineList.appendChild(timelineFrag);
    }
  }

  // ==========================================================================
  // 9. ACTIONS & MUTATIONS (OPTIMISTIC INSTANT UPDATES)
  // ==========================================================================
  function selectTask(taskId, isUserClick) {
    var isDifferentTask = (state.selectedTaskId !== taskId);
    state.selectedTaskId = taskId;
    var rows = el.taskItemsList.querySelectorAll('.task-item-row');
    rows.forEach(function (r) {
      if (r.dataset.taskId === taskId) {
        r.classList.add('selected');
        if (isUserClick) {
          r.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
      } else {
        r.classList.remove('selected');
      }
    });

    renderTaskDetails(isDifferentTask);
  }

  function setPersonFilter(memberName) {
    state.filterType = 'PERSON';
    state.selectedFilter = memberName;
    renderAll();
  }

  function setViewFilter(filterType) {
    state.filterType = filterType;
    state.selectedFilter = filterType;
    renderAll();
  }

  function syncCustomStatusSelect() {
    if (!el.updateStatusTrigger || !el.updateStatusMenu || !el.updateStatusSelect) return;

    var currentVal = el.updateStatusSelect.value || (DataService.statuses && DataService.statuses[0]) || 'Register';
    if (el.updateStatusLabel) {
      el.updateStatusLabel.textContent = currentVal;
      if (isKhmer(currentVal)) {
        el.updateStatusLabel.setAttribute('lang', 'km');
      } else {
        el.updateStatusLabel.removeAttribute('lang');
      }
    }

    el.updateStatusMenu.innerHTML = '';
    (DataService.statuses || []).forEach(function (st) {
      var optBtn = document.createElement('button');
      optBtn.type = 'button';
      var isSelected = (st === currentVal);
      optBtn.className = 'custom-select-option' + (isSelected ? ' is-selected' : '');
      optBtn.setAttribute('role', 'option');
      optBtn.setAttribute('aria-selected', isSelected ? 'true' : 'false');
      optBtn.dataset.value = st;

      var textSpan = document.createElement('span');
      textSpan.className = 'custom-select-option-text';
      textSpan.textContent = st;
      if (isKhmer(st)) textSpan.setAttribute('lang', 'km');
      optBtn.appendChild(textSpan);

      if (isSelected) {
        var checkSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        checkSvg.setAttribute('class', 'custom-select-check');
        checkSvg.setAttribute('viewBox', '0 0 24 24');
        checkSvg.setAttribute('fill', 'none');
        checkSvg.setAttribute('stroke', 'currentColor');
        checkSvg.setAttribute('stroke-width', '2.5');
        checkSvg.innerHTML = '<polyline points="20 6 9 17 4 12"/>';
        optBtn.appendChild(checkSvg);
      }

      optBtn.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        el.updateStatusSelect.value = st;
        closeCustomStatusMenu();
        syncCustomStatusSelect();
      });

      el.updateStatusMenu.appendChild(optBtn);
    });
  }

  function toggleCustomStatusMenu() {
    if (!el.updateStatusMenu) return;
    if (el.updateStatusMenu.classList.contains('hidden')) {
      openCustomStatusMenu();
    } else {
      closeCustomStatusMenu();
    }
  }

  function openCustomStatusMenu() {
    if (!el.updateStatusMenu || !el.updateStatusTrigger) return;
    syncCustomStatusSelect();

    // Check if dropdown will be cut off or blocked below
    var triggerRect = el.updateStatusTrigger.getBoundingClientRect();
    var estimatedHeight = Math.min(220, (DataService.statuses ? DataService.statuses.length : 4) * 34 + 16);
    var spaceBelow = window.innerHeight - triggerRect.bottom;
    var spaceAbove = triggerRect.top;

    var scrollParent = el.updateStatusTrigger.closest('.details-scroll-container') ||
                       el.updateStatusTrigger.closest('#panel-details') ||
                       el.updateStatusTrigger.closest('.details-content');
    if (scrollParent) {
      var parentRect = scrollParent.getBoundingClientRect();
      spaceBelow = Math.min(spaceBelow, parentRect.bottom - triggerRect.bottom);
    }

    if (spaceBelow < estimatedHeight && spaceAbove > spaceBelow) {
      el.updateStatusMenu.classList.add('drop-up');
    } else {
      el.updateStatusMenu.classList.remove('drop-up');
    }

    el.updateStatusMenu.classList.remove('hidden');
    el.updateStatusTrigger.classList.add('is-open');
    el.updateStatusTrigger.setAttribute('aria-expanded', 'true');
  }

  function closeCustomStatusMenu() {
    if (!el.updateStatusMenu || !el.updateStatusTrigger) return;
    el.updateStatusMenu.classList.add('hidden');
    el.updateStatusMenu.classList.remove('drop-up');
    el.updateStatusTrigger.classList.remove('is-open');
    el.updateStatusTrigger.setAttribute('aria-expanded', 'false');
  }

  var confirmModalResolver = null;

  function showConfirmModal(options) {
    options = options || {};
    if (confirmModalResolver) {
      confirmModalResolver(false);
      confirmModalResolver = null;
    }

    if (!el.confirmModal) {
      return Promise.resolve(true);
    }

    el.confirmModalTitle.textContent = options.title || 'Confirmation';
    el.confirmModalHeadline.textContent = options.headline || 'Are you sure?';
    el.confirmModalMessage.textContent = options.message || '';

    // Icon
    var wrapClass = 'confirm-modal-icon-wrap';
    var iconHtml = '';
    if (options.iconType === 'check') {
      wrapClass += ' is-success';
      iconHtml = '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>';
    } else if (options.iconType === 'archive') {
      wrapClass += ' is-warning';
      iconHtml = '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="21 8 21 21 3 21 3 8"/><rect x="1" y="3" width="22" height="5"/><line x1="10" y1="12" x2="14" y2="12"/></svg>';
    } else if (options.iconType === 'restore') {
      wrapClass += ' is-accent';
      iconHtml = '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>';
    } else if (options.iconType === 'danger') {
      wrapClass += ' is-danger';
      iconHtml = '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/></svg>';
    } else {
      wrapClass += (options.confirmType === 'danger' ? ' is-danger' : ' is-accent');
      iconHtml = '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
    }
    el.confirmModalIconWrap.className = wrapClass;
    el.confirmModalIconWrap.innerHTML = iconHtml;

    // Buttons
    el.btnOkConfirm.textContent = options.confirmText || 'Confirm';
    if (options.confirmType === 'danger') {
      el.btnOkConfirm.className = 'btn btn-danger';
    } else {
      el.btnOkConfirm.className = 'btn btn-primary';
    }

    if (options.isAlert) {
      el.btnCancelConfirm.classList.add('hidden');
    } else {
      el.btnCancelConfirm.classList.remove('hidden');
      el.btnCancelConfirm.textContent = options.cancelText || 'Cancel';
    }

    el.confirmModal.classList.remove('hidden');

    // Ensure buttons on the same page/modal share equal size with zero text cut-off
    if (el.btnCancelConfirm && el.btnOkConfirm) {
      el.btnCancelConfirm.style.width = 'auto';
      el.btnOkConfirm.style.width = 'auto';
      var wCancel = options.isAlert ? 0 : el.btnCancelConfirm.scrollWidth;
      var wOk = el.btnOkConfirm.scrollWidth;
      var equalW = Math.max(wCancel, wOk, 170);
      if (!options.isAlert) {
        el.btnCancelConfirm.style.width = equalW + 'px';
      }
      el.btnOkConfirm.style.width = equalW + 'px';
    }

    el.btnOkConfirm.focus();

    return new Promise(function (resolve) {
      confirmModalResolver = resolve;
    });
  }

  function closeConfirmModal(result) {
    if (el.confirmModal) {
      el.confirmModal.classList.add('hidden');
    }
    if (el.btnCancelConfirm) el.btnCancelConfirm.style.width = '';
    if (el.btnOkConfirm) el.btnOkConfirm.style.width = '';
    if (confirmModalResolver) {
      var r = confirmModalResolver;
      confirmModalResolver = null;
      r(Boolean(result));
    }
  }

  function showAlertModal(message, title) {
    return showConfirmModal({
      title: title || 'Notice',
      headline: title || 'Notice',
      message: message,
      confirmText: 'OK',
      isAlert: true
    });
  }

  async function confirmTaskStateChange(taskId, targetState) {
    var task = DataService.getTaskById(taskId);
    if (!task) return;

    var title, headline, message, confirmText, confirmType, iconType;
    var taskTitle = task.taskName ? ('"' + task.taskName + '"') : 'this task';

    if (targetState === 'Completed') {
      title = 'Complete Task';
      headline = 'Mark task as completed?';
      message = 'Are you sure you want to mark ' + taskTitle + ' as completed?';
      confirmText = 'Complete';
      confirmType = 'primary';
      iconType = 'check';
    } else if (targetState === 'Archived') {
      title = 'Delete Task';
      headline = 'Move to Deleted tasks?';
      message = 'Are you sure you want to move ' + taskTitle + ' to Deleted tasks? You can restore it later.';
      confirmText = 'Delete';
      confirmType = 'danger';
      iconType = 'danger';
    } else {
      title = 'Restore Task';
      headline = 'Restore task to Active?';
      message = 'Are you sure you want to restore ' + taskTitle + ' back to active tasks?';
      confirmText = 'Restore';
      confirmType = 'primary';
      iconType = 'restore';
    }

    var confirmed = await showConfirmModal({
      title: title,
      headline: headline,
      message: message,
      confirmText: confirmText,
      cancelText: 'Cancel',
      confirmType: confirmType,
      iconType: iconType
    });

    if (confirmed) {
      handleTaskStateChange(taskId, targetState);
    }
  }

  async function handleStatusUpdate(e) {
    if (e && e.preventDefault) e.preventDefault();
    if (!state.selectedTaskId || !state.currentUser) return;
    if (state.isUpdatingStatus) return;

    var taskId = state.selectedTaskId;
    var newStatus = el.updateStatusSelect.value;
    var note = el.updateNoteInput.value.trim();

    var task = DataService.getTaskById(taskId);
    if (!task) return;

    // Guard: ignore if status is identical and note is empty
    if (task.status === newStatus && !note) {
      return;
    }

    state.isUpdatingStatus = true;
    if (el.btnSubmitStatusUpdate) {
      el.btnSubmitStatusUpdate.disabled = true;
      el.btnSubmitStatusUpdate.textContent = 'Updating…';
    }

    var prevStatus = task.status;
    var prevLastUpdatedBy = task.lastUpdatedBy;
    var prevLastUpdatedAt = task.lastUpdatedAt;

    var nowIso = new Date().toISOString();
    var historyId = 'HIST-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();

    task.status = newStatus;
    task.lastUpdatedBy = state.currentUser.name;
    task.lastUpdatedAt = nowIso;

    var historyEntry = {
      historyId: historyId,
      taskId: taskId,
      taskName: task.taskName,
      status: newStatus,
      note: note,
      updatedBy: state.currentUser.name,
      timestamp: nowIso
    };
    DataService.history.unshift(historyEntry);
    DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);

    // Instant UI update
    el.updateNoteInput.value = '';
    state.lastRenderedTaskId = null;
    renderAll(false);
    selectTask(taskId, false);
    DataService.saveToStorage();
    updateSyncStatus('saving');

    setTimeout(function () {
      state.isUpdatingStatus = false;
      if (el.btnSubmitStatusUpdate) {
        el.btnSubmitStatusUpdate.disabled = false;
        el.btnSubmitStatusUpdate.textContent = 'Update Status';
      }
    }, 500);

    var rollbackStatus = function () {
      task.status = prevStatus;
      task.lastUpdatedBy = prevLastUpdatedBy;
      task.lastUpdatedAt = prevLastUpdatedAt;
      DataService.history = DataService.history.filter(function (h) {
        return String(h.historyId || '').trim() !== String(historyId).trim();
      });
      renderAll();
      selectTask(taskId);
    };

    // Sync in background with retry and idempotency
    executeUserActionWithRetry({
      action: 'updateTask',
      taskName: task.taskName,
      payload: {
        taskId: taskId,
        status: newStatus,
        note: note,
        historyId: historyId,
        timestamp: nowIso
      },
      rollbackFn: rollbackStatus,
      onSuccess: function (res) {
        if (res.task) {
          var updated = res.task;
          var idx = DataService.tasks.findIndex(function (t) {
            return String(t.taskId || '').trim().toLowerCase() === String(taskId).trim().toLowerCase();
          });
          if (idx !== -1) DataService.tasks[idx] = updated;
          if (updated.lastUpdatedAt) DataService.lastSync = updated.lastUpdatedAt;
        }
        if (res.historyEntry) {
          var hIdx = DataService.history.findIndex(function (h) {
            return String(h.historyId || '').trim() === String(historyId).trim();
          });
          if (hIdx !== -1) {
            DataService.history[hIdx] = res.historyEntry;
          } else {
            DataService.history.unshift(res.historyEntry);
          }
        }
        DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);
        DataService.saveToStorage();
        renderTaskDetails(false);
      }
    });
  }

  function handleTaskStateChange(taskId, targetState) {
    var task = DataService.getTaskById(taskId);
    if (!task || !state.currentUser) return;
    if (task.state === targetState) return;
    if (state.pendingStateTasks.has(taskId)) return;

    state.pendingStateTasks.add(taskId);
    setTimeout(function () {
      state.pendingStateTasks.delete(taskId);
    }, 600);

    var prevState = task.state || 'Active';
    var prevLastUpdatedBy = task.lastUpdatedBy;
    var prevLastUpdatedAt = task.lastUpdatedAt;

    var nowIso = new Date().toISOString();
    var historyId = 'HIST-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();

    task.state = targetState;
    task.lastUpdatedBy = state.currentUser.name;
    task.lastUpdatedAt = nowIso;

    var historyEntry = {
      historyId: historyId,
      taskId: taskId,
      taskName: task.taskName,
      status: targetState === 'Completed' ? 'Done' : task.status,
      note: 'State marked as ' + targetState,
      updatedBy: state.currentUser.name,
      timestamp: nowIso
    };
    DataService.history.unshift(historyEntry);
    DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);

    // Show Undo Toast
    showUndoToast(taskId, prevState, targetState);

    // Calculate next visible task to auto-advance selection
    var currentFiltered = getFilteredTasks();
    var currentIndex = currentFiltered.findIndex(function (t) { return t.taskId === taskId; });
    var nextTaskId = null;
    if (currentIndex !== -1) {
      if (currentIndex + 1 < currentFiltered.length) {
        nextTaskId = currentFiltered[currentIndex + 1].taskId;
      } else if (currentIndex - 1 >= 0) {
        nextTaskId = currentFiltered[currentIndex - 1].taskId;
      }
    }

    // Instant UI update
    renderAll();
    DataService.saveToStorage();
    updateSyncStatus('saving');

    // Auto-advance selection if archived task is removed from current view
    var remainingVisible = getFilteredTasks();
    var isStillVisible = remainingVisible.some(function (t) { return t.taskId === taskId; });
    if (isStillVisible) {
      selectTask(taskId);
    } else if (nextTaskId && remainingVisible.some(function (t) { return t.taskId === nextTaskId; })) {
      selectTask(nextTaskId);
    } else if (remainingVisible.length > 0) {
      selectTask(remainingVisible[0].taskId);
    } else {
      state.selectedTaskId = null;
      state.lastRenderedTaskId = null;
      renderTaskDetails(true);
    }

    var rollbackState = function () {
      task.state = prevState;
      task.lastUpdatedBy = prevLastUpdatedBy;
      task.lastUpdatedAt = prevLastUpdatedAt;
      DataService.history = DataService.history.filter(function (h) {
        return String(h.historyId || '').trim() !== String(historyId).trim();
      });
      renderAll();
      selectTask(taskId);
    };

    // Background sync with retry and idempotency
    executeUserActionWithRetry({
      action: 'setState',
      taskName: task.taskName,
      payload: {
        taskId: taskId,
        state: targetState,
        note: 'State set to ' + targetState,
        historyId: historyId,
        timestamp: nowIso
      },
      rollbackFn: rollbackState,
      onSuccess: function (res) {
        if (res.task && res.task.lastUpdatedAt) {
          DataService.lastSync = res.task.lastUpdatedAt;
        }
        if (res.historyEntry) {
          var hIdx = DataService.history.findIndex(function (h) {
            return String(h.historyId || '').trim() === String(historyId).trim();
          });
          if (hIdx !== -1) {
            DataService.history[hIdx] = res.historyEntry;
          }
        }
        DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);
        DataService.saveToStorage();
      }
    });
  }

  async function confirmAndDeleteTask(taskId) {
    var task = DataService.getTaskById(taskId);
    if (!task || !state.currentUser) return;

    var confirmed = await showConfirmModal({
      title: 'Permanently Delete Task',
      headline: 'Permanently delete "' + task.taskName + '"?',
      message: 'Are you sure you want to permanently delete this task? This action cannot be undone and will remove all history.',
      confirmText: 'Delete Permanently',
      cancelText: 'Cancel',
      confirmType: 'danger',
      iconType: 'danger'
    });

    if (!confirmed) return;

    var prevTasks = DataService.tasks.slice();
    var prevHistory = DataService.history.slice();
    var prevSelectedId = state.selectedTaskId;

    // Instant optimistic removal
    DataService.tasks = DataService.tasks.filter(function (t) {
      return String(t.taskId || '').toLowerCase() !== String(taskId).toLowerCase();
    });
    DataService.history = DataService.history.filter(function (h) {
      return String(h.taskId || '').toLowerCase() !== String(taskId).toLowerCase();
    });

    if (state.selectedTaskId === taskId) {
      var remaining = getFilteredTasks();
      state.selectedTaskId = remaining.length > 0 ? remaining[0].taskId : null;
      state.lastRenderedTaskId = null;
    }

    renderAll();
    DataService.saveToStorage();
    updateSyncStatus('saving');

    var rollbackDelete = function () {
      DataService.tasks = prevTasks;
      DataService.history = prevHistory;
      state.selectedTaskId = prevSelectedId;
      renderAll();
      if (prevSelectedId) selectTask(prevSelectedId);
    };

    // Delete in background with retry and idempotency
    executeUserActionWithRetry({
      action: 'deleteTask',
      taskName: taskId,
      payload: { taskId: taskId },
      rollbackFn: rollbackDelete,
      onSuccess: function () {
        DataService.saveToStorage();
      }
    });
  }

  function showUndoToast(taskId, prevState, newState) {
    if (state.undoTimeout) clearTimeout(state.undoTimeout);
    state.lastUndoAction = { taskId: taskId, prevState: prevState };

    var actionText = newState === 'Archived' ? 'moved to Deleted' : (newState === 'Completed' ? 'marked as Completed' : 'restored to Active');
    el.undoToastMessage.textContent = 'Task ' + actionText;
    el.undoToast.classList.remove('hidden');

    state.undoTimeout = setTimeout(function () {
      el.undoToast.classList.add('hidden');
      state.lastUndoAction = null;
    }, 6000);
  }

  function handleUndoToastClick() {
    if (!state.lastUndoAction) return;
    var action = state.lastUndoAction;
    el.undoToast.classList.add('hidden');
    handleTaskStateChange(action.taskId, action.prevState);
    state.lastUndoAction = null;
  }

  function handleTogglePrivacy() {
    if (!state.selectedTaskId) return;
    var task = DataService.getTaskById(state.selectedTaskId);
    if (!task) return;

    var currentUserName = (state.currentUser ? state.currentUser.name : '').trim().toLowerCase();
    var isCreator = task.createdBy && task.createdBy.trim().toLowerCase() === currentUserName;
    if (!isCreator) return;

    var newIsPrivate = !el.detailShareTeamCheckbox.checked;
    var prevIsPrivate = task.isPrivate;
    task.isPrivate = newIsPrivate;

    // Instant optimistic UI update
    if (el.detailPrivacyBadge) {
      if (task.isPrivate) {
        el.detailPrivacyBadge.classList.remove('hidden');
      } else {
        el.detailPrivacyBadge.classList.add('hidden');
      }
    }
    if (el.detailPrivacyToggleWrapper) {
      if (newIsPrivate) {
        el.detailPrivacyToggleWrapper.classList.add('is-private');
        el.detailPrivacyToggleWrapper.classList.remove('is-shared');
      } else {
        el.detailPrivacyToggleWrapper.classList.add('is-shared');
        el.detailPrivacyToggleWrapper.classList.remove('is-private');
      }
    }
    DataService.saveToStorage();
    renderPeoplePanel();
    renderTasksList();

    executeUserActionWithRetry({
      action: 'setPrivacy',
      taskName: task.taskName,
      payload: {
        taskId: task.taskId,
        isPrivate: newIsPrivate
      },
      rollbackFn: function () {
        task.isPrivate = prevIsPrivate;
        renderAll();
      },
      onSuccess: function (res) {
        if (res && res.task) {
          var idx = DataService.tasks.findIndex(function (t) {
            return String(t.taskId || '').trim().toLowerCase() === String(task.taskId || '').trim().toLowerCase();
          });
          if (idx !== -1) {
            DataService.tasks[idx].isPrivate = (res.task.isPrivate === true);
          }
        }
        // Public/Private toggle is silent: no Activity History entry created
        DataService.saveToStorage();
        renderAll();
      }
    });
  }

  // ==========================================================================
  // 10. MODAL: CREATE TASK
  // ==========================================================================
  function updateDeadlineHint(dateVal) {
    if (!el.taskDeadlineHint) return;
    if (!dateVal) {
      el.taskDeadlineHint.textContent = '';
      return;
    }
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var selected = new Date(dateVal.includes('T') ? dateVal : (dateVal + 'T00:00:00'));
    if (isNaN(selected.getTime())) {
      el.taskDeadlineHint.textContent = '';
      return;
    }
    var diffMs = selected.getTime() - today.getTime();
    var diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));

    if (diffDays === 14) {
      el.taskDeadlineHint.textContent = '(2 weeks from now selected)';
    } else if (diffDays === 1) {
      el.taskDeadlineHint.textContent = '(1 day from now)';
    } else if (diffDays > 1) {
      el.taskDeadlineHint.textContent = '(' + diffDays + ' days from now)';
    } else if (diffDays === 0) {
      el.taskDeadlineHint.textContent = '(Today)';
    } else if (diffDays === -1) {
      el.taskDeadlineHint.textContent = '(1 day ago)';
    } else {
      el.taskDeadlineHint.textContent = '(' + Math.abs(diffDays) + ' days ago)';
    }
  }

  function openNewTaskModal() {
    state.editingTaskId = null;
    el.newTaskForm.reset();
    el.taskNameError.textContent = '';
    el.taskAssigneeError.textContent = '';
    el.taskDeadlineError.textContent = '';
    if (el.taskDescError) el.taskDescError.textContent = '';
    el.taskNameInput.classList.remove('is-invalid');
    el.taskAssigneeInput.classList.remove('is-invalid');
    el.taskDeadlineInput.classList.remove('is-invalid');
    el.taskDescInput.classList.remove('is-invalid');

    if (el.modalHeading) el.modalHeading.textContent = 'Create New Task';
    if (el.btnSaveTask) el.btnSaveTask.textContent = 'Create Task';
    if (el.taskEditStatusGroup) el.taskEditStatusGroup.classList.add('hidden');
    if (el.taskModalInfoCallout) el.taskModalInfoCallout.classList.remove('hidden');

    el.taskNameInput.value = '';
    el.taskDescInput.value = '';

    // Populate Assignee Select with Active Users
    el.taskAssigneeInput.innerHTML = '';
    var currentUserName = (state.currentUser && state.currentUser.name) ? state.currentUser.name : '';

    if (DataService.users.length > 0) {
      DataService.users.forEach(function (user) {
        var opt = document.createElement('option');
        opt.value = user.name;
        opt.textContent = user.name;
        if (user.name.toLowerCase() === currentUserName.toLowerCase()) {
          opt.selected = true;
        }
        el.taskAssigneeInput.appendChild(opt);
      });
    } else if (currentUserName) {
      var optSelf = document.createElement('option');
      optSelf.value = currentUserName;
      optSelf.textContent = currentUserName;
      optSelf.selected = true;
      el.taskAssigneeInput.appendChild(optSelf);
    }

    if (currentUserName) {
      el.taskAssigneeInput.value = currentUserName;
    }
    // Auto-select deadline 2 weeks (+14 days) from current date
    var defDate = new Date();
    defDate.setDate(defDate.getDate() + 14);
    var yyyy = defDate.getFullYear();
    var mm = String(defDate.getMonth() + 1).padStart(2, '0');
    var dd = String(defDate.getDate()).padStart(2, '0');
    el.taskDeadlineInput.value = `${yyyy}-${mm}-${dd}`;
    updateDeadlineHint(el.taskDeadlineInput.value);

    // Initial assignee tracking and UX default for "Share with team"
    var creatorName = currentUserName.trim().toLowerCase();
    var initialAssignee = (el.taskAssigneeInput.value || '').trim();
    var initialAssigneeKey = initialAssignee.toLowerCase();
    state.modalAssigneePrevious = initialAssignee;

    var shouldShare = Boolean(initialAssigneeKey && initialAssigneeKey !== creatorName);
    if (el.taskShareTeamInput) {
      el.taskShareTeamInput.checked = shouldShare;
    }
    updateNewTaskModalPrivacy(shouldShare);

    el.newTaskModal.classList.remove('hidden');
    el.taskNameInput.focus();
  }

  function openEditTaskModal(taskId) {
    var task = DataService.getTaskById(taskId);
    if (!task) return;

    var currentUserName = (state.currentUser && state.currentUser.name) ? state.currentUser.name.trim().toLowerCase() : '';
    var isCreator = task.createdBy && task.createdBy.trim().toLowerCase() === currentUserName;
    if (!isCreator) return;

    state.editingTaskId = task.taskId;
    el.newTaskForm.reset();
    el.taskNameError.textContent = '';
    el.taskAssigneeError.textContent = '';
    el.taskDeadlineError.textContent = '';
    if (el.taskDescError) el.taskDescError.textContent = '';
    el.taskNameInput.classList.remove('is-invalid');
    el.taskAssigneeInput.classList.remove('is-invalid');
    el.taskDeadlineInput.classList.remove('is-invalid');
    el.taskDescInput.classList.remove('is-invalid');

    if (el.modalHeading) el.modalHeading.textContent = 'Edit Task';
    if (el.btnSaveTask) el.btnSaveTask.textContent = 'Save Changes';
    if (el.taskEditStatusGroup) el.taskEditStatusGroup.classList.remove('hidden');
    if (el.taskModalInfoCallout) el.taskModalInfoCallout.classList.add('hidden');

    el.taskNameInput.value = task.taskName || '';
    el.taskDescInput.value = task.description || '';
    el.taskDeadlineInput.value = task.deadline || '';
    updateDeadlineHint(el.taskDeadlineInput.value);

    // Populate Assignees and select current assignee
    el.taskAssigneeInput.innerHTML = '';
    if (DataService.users.length > 0) {
      DataService.users.forEach(function (user) {
        var opt = document.createElement('option');
        opt.value = user.name;
        opt.textContent = user.name;
        if (user.name.toLowerCase() === (task.assignedTo || '').toLowerCase()) {
          opt.selected = true;
        }
        el.taskAssigneeInput.appendChild(opt);
      });
    } else if (task.assignedTo) {
      var optA = document.createElement('option');
      optA.value = task.assignedTo;
      optA.textContent = task.assignedTo;
      optA.selected = true;
      el.taskAssigneeInput.appendChild(optA);
    }
    el.taskAssigneeInput.value = task.assignedTo || '';

    // Populate Statuses and select current status
    if (el.taskStatusInput) {
      el.taskStatusInput.innerHTML = '';
      DataService.statuses.forEach(function (st) {
        var optSt = document.createElement('option');
        optSt.value = st;
        optSt.textContent = st;
        if (st.toLowerCase() === (task.status || '').toLowerCase()) {
          optSt.selected = true;
        }
        el.taskStatusInput.appendChild(optSt);
      });
      el.taskStatusInput.value = task.status || (DataService.statuses[0] || 'Register');
    }

    if (el.taskShareTeamInput) {
      el.taskShareTeamInput.checked = !task.isPrivate;
    }
    state.modalAssigneePrevious = task.assignedTo || '';
    updateNewTaskModalPrivacy(!task.isPrivate);

    el.newTaskModal.classList.remove('hidden');
    el.taskNameInput.focus();
  }

  function updateNewTaskModalPrivacy(isShared) {
    var modalOverlay = el.newTaskModal;
    var modalDialog = document.getElementById('new-task-modal-dialog') || (modalOverlay ? modalOverlay.querySelector('.modal-dialog') : null);
    if (modalOverlay) {
      if (isShared) {
        modalOverlay.classList.add('is-shared');
        modalOverlay.classList.remove('is-private');
      } else {
        modalOverlay.classList.add('is-private');
        modalOverlay.classList.remove('is-shared');
      }
    }
    if (modalDialog) {
      if (isShared) {
        modalDialog.classList.add('is-shared');
        modalDialog.classList.remove('is-private');
      } else {
        modalDialog.classList.add('is-private');
        modalDialog.classList.remove('is-shared');
      }
    }
  }

  function closeNewTaskModal() {
    el.newTaskModal.classList.add('hidden');
    state.modalAssigneePrevious = null;
    state.editingTaskId = null;
    if (el.taskEditStatusGroup) el.taskEditStatusGroup.classList.add('hidden');
    if (el.taskModalInfoCallout) el.taskModalInfoCallout.classList.remove('hidden');
    if (el.modalHeading) el.modalHeading.textContent = 'Create New Task';
    if (el.btnSaveTask) el.btnSaveTask.textContent = 'Create Task';
  }

  async function handleCreateTaskSubmit(e) {
    e.preventDefault();
    if (state.isCreatingTask) return;

    if (state.editingTaskId) {
      handleEditTaskSubmit(e);
      return;
    }

    el.taskNameError.textContent = '';
    el.taskAssigneeError.textContent = '';
    el.taskDeadlineError.textContent = '';
    if (el.taskDescError) el.taskDescError.textContent = '';

    var name = el.taskNameInput.value.trim();
    var assignedTo = el.taskAssigneeInput.value.trim();
    var deadline = el.taskDeadlineInput.value;
    var description = el.taskDescInput.value.trim();

    var hasError = false;

    if (!name) {
      el.taskNameError.textContent = 'Task name is required';
      el.taskNameInput.classList.add('is-invalid');
      hasError = true;
    } else {
      el.taskNameInput.classList.remove('is-invalid');
    }

    if (!assignedTo) {
      el.taskAssigneeError.textContent = 'Assigned To is required';
      el.taskAssigneeInput.classList.add('is-invalid');
      hasError = true;
    } else {
      el.taskAssigneeInput.classList.remove('is-invalid');
    }

    if (!deadline) {
      el.taskDeadlineError.textContent = 'Deadline is required';
      el.taskDeadlineInput.classList.add('is-invalid');
      hasError = true;
    } else {
      el.taskDeadlineInput.classList.remove('is-invalid');
    }

    if (!description) {
      if (el.taskDescError) el.taskDescError.textContent = 'Description is required';
      el.taskDescInput.classList.add('is-invalid');
      hasError = true;
    } else {
      el.taskDescInput.classList.remove('is-invalid');
    }

    if (hasError) return;

    state.isCreatingTask = true;
    if (el.btnSaveTask) {
      el.btnSaveTask.disabled = true;
      el.btnSaveTask.textContent = 'Creating…';
    }

    // 1. Instant optimistic creation
    var nowIso = new Date().toISOString();
    var tempId = 'TSK-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();
    var taskCode = generateNewTaskCode(DataService.tasks, new Date());
    var tempHistoryId = 'HIST-' + Date.now().toString(36).toUpperCase() + '-INIT';
    var currentUserName = state.currentUser ? state.currentUser.name : assignedTo;
    var isPrivate = el.taskShareTeamInput ? !el.taskShareTeamInput.checked : true;

    var optimisticTask = {
      taskId: tempId,
      taskCode: taskCode,
      taskName: name,
      description: description,
      assignedTo: assignedTo,
      deadline: deadline,
      status: (DataService.statuses && DataService.statuses.length > 0) ? DataService.statuses[0] : 'Register',
      state: 'Active',
      createdBy: currentUserName,
      createdAt: nowIso,
      lastUpdatedBy: currentUserName,
      lastUpdatedAt: nowIso,
      isPrivate: isPrivate
    };

    var optimisticHistory = {
      historyId: tempHistoryId,
      taskId: tempId,
      taskName: name,
      status: (DataService.statuses && DataService.statuses.length > 0) ? DataService.statuses[0] : 'Register',
      note: 'Task created',
      updatedBy: currentUserName,
      timestamp: nowIso
    };

    DataService.tasks.unshift(optimisticTask);
    DataService.history.unshift(optimisticHistory);
    DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);

    // Close modal and render instantly without waiting for backend
    closeNewTaskModal();
    renderAll();
    selectTask(tempId);
    DataService.saveToStorage();
    updateSyncStatus('saving');

    setTimeout(function () {
      state.isCreatingTask = false;
      if (el.btnSaveTask) {
        el.btnSaveTask.disabled = false;
        el.btnSaveTask.textContent = 'Create Task';
      }
    }, 400);

    var rollbackCreate = function () {
      DataService.tasks = DataService.tasks.filter(function (t) {
        return String(t.taskId || '').trim().toLowerCase() !== String(tempId).trim().toLowerCase();
      });
      DataService.history = DataService.history.filter(function (h) {
        return String(h.taskId || '').trim().toLowerCase() !== String(tempId).trim().toLowerCase();
      });
      if (state.selectedTaskId === tempId) state.selectedTaskId = null;
      renderAll();
    };

    // 2. Perform backend request in the background with retry and idempotency
    executeUserActionWithRetry({
      action: 'createTask',
      taskName: name,
      payload: {
        taskId: tempId,
        taskCode: taskCode,
        historyId: tempHistoryId,
        taskName: name,
        assignedTo: assignedTo,
        deadline: deadline,
        description: description,
        isPrivate: isPrivate
      },
      rollbackFn: rollbackCreate,
      onSuccess: function (res) {
        if (res && res.task) {
          var realTask = res.task;
          var realHist = res.historyEntry;
          var idx = DataService.tasks.findIndex(function (t) {
            return String(t.taskId || '').trim().toLowerCase() === String(tempId).trim().toLowerCase();
          });
          if (idx !== -1) {
            DataService.tasks[idx] = realTask;
          }
          DataService.history.forEach(function (h) {
            if (String(h.taskId || '').trim().toLowerCase() === String(tempId).trim().toLowerCase()) {
              h.taskId = realTask.taskId;
              if (realHist && h.historyId === tempHistoryId) {
                h.historyId = realHist.historyId;
                h.timestamp = realHist.timestamp;
              }
            }
          });
          DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);
          if (state.selectedTaskId === tempId) {
            state.selectedTaskId = realTask.taskId;
          }
          if (realTask.lastUpdatedAt) {
            DataService.lastSync = realTask.lastUpdatedAt;
          }
          DataService.saveToStorage();
          renderAll();
          selectTask(realTask.taskId);
        }
      }
    });
  }

  async function handleEditTaskSubmit(e) {
    e.preventDefault();
    if (state.isCreatingTask) return;

    var taskId = state.editingTaskId;
    var task = DataService.getTaskById(taskId);
    if (!task) {
      closeNewTaskModal();
      return;
    }

    el.taskNameError.textContent = '';
    el.taskAssigneeError.textContent = '';
    el.taskDeadlineError.textContent = '';
    if (el.taskDescError) el.taskDescError.textContent = '';

    var newName = el.taskNameInput.value.trim();
    var newAssignedTo = el.taskAssigneeInput.value.trim();
    var newDeadline = el.taskDeadlineInput.value;
    var newDesc = el.taskDescInput.value.trim();
    var newStatus = el.taskStatusInput ? el.taskStatusInput.value : task.status;
    var newIsPrivate = el.taskShareTeamInput ? !el.taskShareTeamInput.checked : task.isPrivate;

    var hasError = false;
    if (!newName) {
      el.taskNameError.textContent = 'Task name is required';
      el.taskNameInput.classList.add('is-invalid');
      hasError = true;
    }
    if (!newAssignedTo) {
      el.taskAssigneeError.textContent = 'Assigned To is required';
      el.taskAssigneeInput.classList.add('is-invalid');
      hasError = true;
    }
    if (!newDeadline) {
      el.taskDeadlineError.textContent = 'Deadline is required';
      el.taskDeadlineInput.classList.add('is-invalid');
      hasError = true;
    }
    if (!newDesc) {
      if (el.taskDescError) el.taskDescError.textContent = 'Description is required';
      el.taskDescInput.classList.add('is-invalid');
      hasError = true;
    }
    if (hasError) return;

    var statusChanged = (newStatus && newStatus.toLowerCase() !== (task.status || '').toLowerCase());
    var assigneeChanged = (newAssignedTo && newAssignedTo.toLowerCase() !== (task.assignedTo || '').toLowerCase());
    var nameChanged = (newName !== task.taskName);
    var descChanged = (newDesc !== task.description);
    var deadlineChanged = (newDeadline !== task.deadline);
    var privacyChanged = (newIsPrivate !== task.isPrivate);

    if (!statusChanged && !assigneeChanged && !nameChanged && !descChanged && !deadlineChanged && !privacyChanged) {
      closeNewTaskModal();
      return;
    }

    state.isCreatingTask = true;
    if (el.btnSaveTask) {
      el.btnSaveTask.disabled = true;
      el.btnSaveTask.textContent = 'Saving…';
    }

    var prevTaskData = {
      taskName: task.taskName,
      description: task.description,
      assignedTo: task.assignedTo,
      deadline: task.deadline,
      status: task.status,
      isPrivate: task.isPrivate,
      lastUpdatedBy: task.lastUpdatedBy,
      lastUpdatedAt: task.lastUpdatedAt
    };
    var prevHistory = [...DataService.history];

    var nowIso = new Date().toISOString();
    var currentUserName = state.currentUser ? state.currentUser.name : 'Team Member';

    // Apply updates optimistically
    task.taskName = newName;
    task.description = newDesc;
    task.assignedTo = newAssignedTo;
    task.deadline = newDeadline;
    task.status = newStatus;
    task.isPrivate = newIsPrivate;
    task.lastUpdatedBy = currentUserName;
    task.lastUpdatedAt = nowIso;

    // Selective Activity History tracking:
    // - Name & Description edits: NEVER log
    // - Privacy toggle: NEVER log
    // - Status change: LOG
    // - Assignee change: LOG
    var optHistoryEntries = [];
    if (statusChanged) {
      var optStatusEntry = {
        historyId: 'HIST-OPT-' + Date.now() + '-S',
        taskId: task.taskId,
        taskName: newName,
        status: newStatus,
        note: 'Status updated to ' + newStatus,
        updatedBy: currentUserName,
        timestamp: nowIso
      };
      optHistoryEntries.push(optStatusEntry);
      DataService.history.unshift(optStatusEntry);
    }
    if (assigneeChanged) {
      var optAssignEntry = {
        historyId: 'HIST-OPT-' + (Date.now() + 1) + '-A',
        taskId: task.taskId,
        taskName: newName,
        status: newStatus,
        note: 'Assigned to ' + newAssignedTo + (prevTaskData.assignedTo ? (' (previously ' + prevTaskData.assignedTo + ')') : ''),
        updatedBy: currentUserName,
        timestamp: nowIso
      };
      optHistoryEntries.push(optAssignEntry);
      DataService.history.unshift(optAssignEntry);
    }
    if (optHistoryEntries.length > 0) {
      DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);
    }

    closeNewTaskModal();
    renderAll();
    DataService.saveToStorage();
    updateSyncStatus('saving');

    setTimeout(function () {
      state.isCreatingTask = false;
      if (el.btnSaveTask) {
        el.btnSaveTask.disabled = false;
        el.btnSaveTask.textContent = 'Create Task';
      }
    }, 400);

    executeUserActionWithRetry({
      action: 'editTask',
      taskName: newName,
      payload: {
        taskId: task.taskId,
        taskName: newName,
        description: newDesc,
        assignedTo: newAssignedTo,
        deadline: newDeadline,
        status: newStatus,
        isPrivate: newIsPrivate
      },
      rollbackFn: function () {
        Object.assign(task, prevTaskData);
        DataService.history = prevHistory;
        renderAll();
      },
      onSuccess: function (res) {
        if (res && res.task) {
          var idx = DataService.tasks.findIndex(function (t) {
            return String(t.taskId || '').toLowerCase() === String(task.taskId || '').toLowerCase();
          });
          if (idx !== -1) {
            DataService.tasks[idx] = res.task;
          }
        }
        if (res && Array.isArray(res.historyEntries) && res.historyEntries.length > 0) {
          res.historyEntries.forEach(function (h) {
            DataService.history.unshift(h);
          });
          DataService.history = DataService.cleanAndDeduplicateHistory(DataService.history);
        }
        DataService.saveToStorage();
        renderAll();
      }
    });
  }

  // ==========================================================================
  // 11. CONTEXT MENU & SHORTCUTS
  // ==========================================================================
  function openContextMenu(e, taskId) {
    state.contextMenuTaskId = taskId;
    var task = DataService.getTaskById(taskId);
    if (!task) return;

    var currentUserName = (state.currentUser ? state.currentUser.name : '').trim().toLowerCase();
    var isCreator = task.createdBy && task.createdBy.trim().toLowerCase() === currentUserName;
    if (el.ctxEditTask) {
      if (isCreator && task.state !== 'Archived') {
        el.ctxEditTask.classList.remove('hidden');
      } else {
        el.ctxEditTask.classList.add('hidden');
      }
    }

    if (task.state === 'Archived') {
      el.ctxMarkComplete.classList.add('hidden');
      if (el.ctxRestoreTask) el.ctxRestoreTask.classList.remove('hidden');
      if (el.ctxDeleteTask) {
        el.ctxDeleteTask.querySelector('span').textContent = 'Delete Permanently';
      }
    } else {
      el.ctxMarkComplete.classList.remove('hidden');
      el.ctxMarkComplete.querySelector('span').textContent =
        task.state === 'Completed' ? 'Mark Active' : 'Mark Complete';
      if (el.ctxRestoreTask) el.ctxRestoreTask.classList.add('hidden');
      if (el.ctxDeleteTask) {
        el.ctxDeleteTask.querySelector('span').textContent = 'Delete Task';
      }
    }

    var x = Math.min(e.clientX, window.innerWidth - 180);
    var y = Math.min(e.clientY, window.innerHeight - 150);

    el.taskContextMenu.style.left = x + 'px';
    el.taskContextMenu.style.top = y + 'px';
    el.taskContextMenu.classList.remove('hidden');
  }

  function closeContextMenu() {
    el.taskContextMenu.classList.add('hidden');
    state.contextMenuTaskId = null;
  }

  function toggleLeftPanel() {
    el.panelPeople.classList.toggle('collapsed');
  }

  function openDisclaimerModal() {
    if (el.disclaimerModal) el.disclaimerModal.classList.remove('hidden');
  }

  function closeDisclaimerModal() {
    if (el.disclaimerModal) el.disclaimerModal.classList.add('hidden');
  }

  // ==========================================================================
  // 12. KEYBOARD SHORTCUTS & EVENT LISTENERS
  // ==========================================================================
  function setupKeyboardShortcuts() {
    document.addEventListener('keydown', function (e) {
      var isModalOpen = !el.newTaskModal.classList.contains('hidden');
      var isCtxOpen = !el.taskContextMenu.classList.contains('hidden');
      var activeTag = (document.activeElement && document.activeElement.tagName) || '';
      var isTyping = activeTag === 'INPUT' || activeTag === 'TEXTAREA' || activeTag === 'SELECT';

      // Escape
      if (e.key === 'Escape') {
        if (el.disclaimerModal && !el.disclaimerModal.classList.contains('hidden')) {
          closeDisclaimerModal();
          e.preventDefault();
          return;
        } else if (el.confirmModal && !el.confirmModal.classList.contains('hidden')) {
          closeConfirmModal(false);
          e.preventDefault();
          return;
        } else if (el.updateStatusMenu && !el.updateStatusMenu.classList.contains('hidden')) {
          closeCustomStatusMenu();
          e.preventDefault();
        } else if (isCtxOpen) {
          closeContextMenu();
          e.preventDefault();
        } else if (isModalOpen) {
          closeNewTaskModal();
          e.preventDefault();
        } else if (document.activeElement === el.taskSearchInput) {
          state.searchQuery = '';
          el.taskSearchInput.value = '';
          el.btnClearSearch.classList.add('hidden');
          var searchKbd = document.querySelector('.search-kbd');
          if (searchKbd) searchKbd.classList.remove('hidden');
          el.taskSearchInput.blur();
          renderTasksList();
          e.preventDefault();
        }
        return;
      }

      // Ctrl+B: Toggle left panel
      if (e.ctrlKey && e.key.toLowerCase() === 'b') {
        e.preventDefault();
        toggleLeftPanel();
        return;
      }

      // Ctrl+N: Create Task
      if ((e.ctrlKey && e.key.toLowerCase() === 'n') || (e.altKey && e.key.toLowerCase() === 'n')) {
        e.preventDefault();
        openNewTaskModal();
        return;
      }

      // / or Ctrl+K: Focus search
      if ((e.key === '/' && !isTyping && !isModalOpen) || (e.ctrlKey && e.key.toLowerCase() === 'k')) {
        e.preventDefault();
        el.taskSearchInput.focus();
        el.taskSearchInput.select();
        return;
      }

      // 'e', 'E', or Delete: Delete / move to Deleted currently selected task
      if (!isTyping && !isModalOpen && state.selectedTaskId && (e.key === 'e' || e.key === 'E' || e.key === 'Delete')) {
        e.preventDefault();
        var selectedTask = DataService.getTaskById(state.selectedTaskId);
        if (selectedTask) {
          if (selectedTask.state === 'Archived') {
            confirmAndDeleteTask(selectedTask.taskId);
          } else {
            confirmTaskStateChange(selectedTask.taskId, 'Archived');
          }
        }
        return;
      }

      // Arrow navigation across tasks
      if (!isTyping && !isModalOpen && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        var visibleTasks = getFilteredTasks();
        if (visibleTasks.length === 0) return;
        var curIdx = visibleTasks.findIndex(function (t) { return t.taskId === state.selectedTaskId; });

        var nextIdx;
        if (e.key === 'ArrowDown') {
          nextIdx = curIdx < visibleTasks.length - 1 ? curIdx + 1 : 0;
        } else {
          nextIdx = curIdx > 0 ? curIdx - 1 : visibleTasks.length - 1;
        }

        e.preventDefault();
        selectTask(visibleTasks[nextIdx].taskId);
        return;
      }

      // Return: Select / open focused task row
      if (e.key === 'Enter' && !isTyping && !isModalOpen) {
        if (document.activeElement && document.activeElement.classList.contains('task-item-row')) {
          var tId = document.activeElement.dataset.taskId;
          if (tId) selectTask(tId);
        }
      }
    });

    // Close context menu on outside click
    document.addEventListener('click', function (e) {
      if (!el.taskContextMenu.contains(e.target)) {
        closeContextMenu();
      }
    });
  }

  var searchDebounceTimer = null;

  function setupEventListeners() {
    el.loginForm.addEventListener('submit', handleLoginSubmit);
    el.btnLogout.addEventListener('click', logout);

    // Startup retry / logout
    el.btnStartupRetry.addEventListener('click', function () {
      startStartupFlow();
    });
    el.btnStartupLogout.addEventListener('click', logout);

    // Left navigation filters
    el.filterAll.addEventListener('click', function () {
      setViewFilter('ALL');
    });
    if (el.filterPersonal) {
      el.filterPersonal.addEventListener('click', function () {
        setViewFilter('PERSONAL');
      });
    }
    el.filterCompleted.addEventListener('click', function () {
      setViewFilter('COMPLETED');
    });
    el.filterArchived.addEventListener('click', function () {
      setViewFilter('ARCHIVED');
    });

    // Search input debounced 200ms
    el.taskSearchInput.addEventListener('input', function (e) {
      var val = e.target.value;
      var searchKbd = document.querySelector('.search-kbd');
      if (val.trim().length > 0) {
        el.btnClearSearch.classList.remove('hidden');
        if (searchKbd) searchKbd.classList.add('hidden');
      } else {
        el.btnClearSearch.classList.add('hidden');
        if (searchKbd) searchKbd.classList.remove('hidden');
      }

      if (searchDebounceTimer) clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(function () {
        state.searchQuery = val;
        renderTasksList();
      }, 200);
    });

    el.btnClearSearch.addEventListener('click', function () {
      state.searchQuery = '';
      el.taskSearchInput.value = '';
      el.btnClearSearch.classList.add('hidden');
      var searchKbd = document.querySelector('.search-kbd');
      if (searchKbd) searchKbd.classList.remove('hidden');
      renderTasksList();
      el.taskSearchInput.focus();
    });

    // Personal Tasks "Show Only Me Tasks" filter
    if (el.filterOnlyMeCheckbox) {
      el.filterOnlyMeCheckbox.addEventListener('change', function (e) {
        state.onlyMeFilter = e.target.checked;
        renderTasksList();
      });
    }

    // Status update: button only, prevent Enter key submission
    el.statusUpdateForm.addEventListener('submit', function (e) {
      e.preventDefault();
    });
    if (el.btnSubmitStatusUpdate) {
      el.btnSubmitStatusUpdate.addEventListener('click', handleStatusUpdate);
    }

    // Task Privacy Toggle
    if (el.detailShareTeamCheckbox) {
      el.detailShareTeamCheckbox.addEventListener('change', handleTogglePrivacy);
    }

    // Custom Status Dropdown events
    if (el.updateStatusTrigger) {
      el.updateStatusTrigger.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        toggleCustomStatusMenu();
      });
    }

    // Close custom dropdown on outside click
    document.addEventListener('click', function (e) {
      if (el.updateStatusWrapper && !el.updateStatusWrapper.contains(e.target)) {
        closeCustomStatusMenu();
      }
    });

    // Confirm / Alert Modal events
    if (el.btnCloseConfirmModal) {
      el.btnCloseConfirmModal.addEventListener('click', function () {
        closeConfirmModal(false);
      });
    }
    if (el.btnCancelConfirm) {
      el.btnCancelConfirm.addEventListener('click', function () {
        closeConfirmModal(false);
      });
    }
    if (el.btnOkConfirm) {
      el.btnOkConfirm.addEventListener('click', function () {
        closeConfirmModal(true);
      });
    }
    if (el.confirmModal) {
      el.confirmModal.addEventListener('click', function (e) {
        if (e.target === el.confirmModal) {
          closeConfirmModal(false);
        }
      });
    }

    // Disclaimer Modal events
    if (el.btnOpenDisclaimer) {
      el.btnOpenDisclaimer.addEventListener('click', openDisclaimerModal);
    }
    if (el.btnCloseDisclaimerModal) {
      el.btnCloseDisclaimerModal.addEventListener('click', closeDisclaimerModal);
    }
    if (el.btnCloseDisclaimerBottom) {
      el.btnCloseDisclaimerBottom.addEventListener('click', closeDisclaimerModal);
    }
    if (el.disclaimerModal) {
      el.disclaimerModal.addEventListener('click', function (e) {
        if (e.target === el.disclaimerModal) {
          closeDisclaimerModal();
        }
      });
    }

    // Mark complete / Restore / Delete actions (with website confirmation popup)
    el.btnMarkComplete.addEventListener('click', function () {
      if (!state.selectedTaskId) return;
      var task = DataService.getTaskById(state.selectedTaskId);
      if (!task) return;
      var targetState = task.state === 'Completed' ? 'Active' : 'Completed';
      confirmTaskStateChange(task.taskId, targetState);
    });

    if (el.btnRestoreTask) {
      el.btnRestoreTask.addEventListener('click', function () {
        if (!state.selectedTaskId) return;
        var task = DataService.getTaskById(state.selectedTaskId);
        if (!task) return;
        confirmTaskStateChange(task.taskId, 'Active');
      });
    }

    if (el.btnDeleteTask) {
      el.btnDeleteTask.addEventListener('click', function () {
        if (!state.selectedTaskId) return;
        var task = DataService.getTaskById(state.selectedTaskId);
        if (!task) return;
        if (task.state === 'Archived') {
          confirmAndDeleteTask(task.taskId);
        } else {
          confirmTaskStateChange(task.taskId, 'Archived');
        }
      });
    }

    // Undo toast button
    el.btnUndoToast.addEventListener('click', handleUndoToastClick);

    // Modal controls
    el.btnOpenNewTask.addEventListener('click', openNewTaskModal);
    el.btnCloseModal.addEventListener('click', closeNewTaskModal);
    el.btnCancelModal.addEventListener('click', closeNewTaskModal);
    el.newTaskForm.addEventListener('submit', handleCreateTaskSubmit);

    function handleTaskAssigneeChange() {
      if (!el.taskAssigneeInput || !el.taskShareTeamInput) return;
      var selectedAssignee = (el.taskAssigneeInput.value || '').trim();
      var selectedAssigneeKey = selectedAssignee.toLowerCase();
      var prevAssigneeKey = (state.modalAssigneePrevious || '').trim().toLowerCase();

      // Only act if the assignee actually changed
      if (selectedAssigneeKey === prevAssigneeKey) return;
      state.modalAssigneePrevious = selectedAssignee;

      var creatorName = (state.currentUser && state.currentUser.name ? state.currentUser.name : '').trim().toLowerCase();

      if (selectedAssigneeKey && selectedAssigneeKey !== creatorName) {
        // Assigned To = someone else -> auto-check "Share with team" (Public)
        el.taskShareTeamInput.checked = true;
        updateNewTaskModalPrivacy(true);
      } else if (selectedAssigneeKey && selectedAssigneeKey === creatorName) {
        // Assigned To = creator themselves -> auto-uncheck "Share with team" (Private)
        el.taskShareTeamInput.checked = false;
        updateNewTaskModalPrivacy(false);
      }
    }

    if (el.taskAssigneeInput) {
      el.taskAssigneeInput.addEventListener('change', handleTaskAssigneeChange);
      el.taskAssigneeInput.addEventListener('input', handleTaskAssigneeChange);
    }

    if (el.taskShareTeamInput) {
      el.taskShareTeamInput.addEventListener('change', function () {
        updateNewTaskModalPrivacy(this.checked);
      });
    }

    if (el.taskDeadlineInput) {
      el.taskDeadlineInput.addEventListener('input', function () {
        updateDeadlineHint(this.value);
      });
      el.taskDeadlineInput.addEventListener('change', function () {
        updateDeadlineHint(this.value);
      });
    }


    // Context menu actions
    el.ctxOpenTask.addEventListener('click', function () {
      if (state.contextMenuTaskId) selectTask(state.contextMenuTaskId);
      closeContextMenu();
    });

    if (el.detailTaskCode) {
      el.detailTaskCode.addEventListener('click', function () {
        var c = this.textContent.trim();
        if (c && c !== '-' && c !== 'Copied!') {
          try {
            navigator.clipboard.writeText(c);
            var self = this;
            var orig = self.textContent;
            self.textContent = 'Copied!';
            setTimeout(function () {
              self.textContent = orig;
            }, 1000);
          } catch (err) {}
        }
      });
    }

    if (el.ctxEditTask) {
      el.ctxEditTask.addEventListener('click', function () {
        if (state.contextMenuTaskId) {
          selectTask(state.contextMenuTaskId);
          openEditTaskModal(state.contextMenuTaskId);
        }
        closeContextMenu();
      });
    }

    el.ctxChangeStatus.addEventListener('click', function () {
      if (state.contextMenuTaskId) {
        selectTask(state.contextMenuTaskId);
        if (el.updateStatusTrigger) {
          el.updateStatusTrigger.focus();
          openCustomStatusMenu();
        } else {
          el.updateStatusSelect.focus();
        }
      }
      closeContextMenu();
    });

    el.ctxMarkComplete.addEventListener('click', function () {
      if (state.contextMenuTaskId) {
        var t = DataService.getTaskById(state.contextMenuTaskId);
        if (t) {
          var targetState = t.state === 'Completed' ? 'Active' : 'Completed';
          confirmTaskStateChange(t.taskId, targetState);
        }
      }
      closeContextMenu();
    });

    if (el.ctxRestoreTask) {
      el.ctxRestoreTask.addEventListener('click', function () {
        if (state.contextMenuTaskId) {
          var t = DataService.getTaskById(state.contextMenuTaskId);
          if (t) {
            confirmTaskStateChange(t.taskId, 'Active');
          }
        }
        closeContextMenu();
      });
    }

    if (el.ctxDeleteTask) {
      el.ctxDeleteTask.addEventListener('click', function () {
        if (state.contextMenuTaskId) {
          var t = DataService.getTaskById(state.contextMenuTaskId);
          if (t) {
            if (t.state === 'Archived') {
              confirmAndDeleteTask(t.taskId);
            } else {
              confirmTaskStateChange(t.taskId, 'Archived');
            }
          }
        }
        closeContextMenu();
      });
    }

    // Dynamic Khmer language tagging on input/textarea
    document.addEventListener('input', function (e) {
      if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA')) {
        if (isKhmer(e.target.value)) {
          e.target.setAttribute('lang', 'km');
        } else if (e.target.getAttribute('lang') === 'km') {
          e.target.removeAttribute('lang');
        }
      }
    });
  }

  // ==========================================================================
  // 13. INITIALIZATION
  // ==========================================================================
  document.addEventListener('DOMContentLoaded', function () {
    setupKeyboardShortcuts();
    setupEventListeners();
    initAuth();
  });

})();
