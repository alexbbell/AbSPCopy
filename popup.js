// popup.js
//
// MUST include .js extensions
import {
  updateStatus,
  switchTab,
  renderVersionTable,
  renderContentTypeTable
} from './ui.js';

import {
  scrapeSharePointColumns,
  provisionSharePointList,
  addSharePointColumns,
  removeNonSystemColumns,
  duplicateSharePointItem,
  fetchSharePointVersions,
  fetchSharePointContentTypes
} from './injectedScripts.js';


function init() {
  const tabBtns = [
    document.getElementById('tab-export'),
    document.getElementById('tab-import'),
    document.getElementById('tab-add-columns'),
    document.getElementById('tab-copy-item'),
    document.getElementById('tab-versions'),
    document.getElementById('tab-ct-manager')
  ];

  const tabContents = [
    document.getElementById('content-export'),
    document.getElementById('content-import'),
    document.getElementById('content-add-columns'),
    document.getElementById('content-copy-item'),
    document.getElementById('content-versions'),
    document.getElementById('content-ct-manager')
  ];

  tabBtns.forEach((btn, index) => {
    if (btn) {
      btn.addEventListener('click', () => switchTab(btn, tabContents[index], tabBtns, tabContents));
    }
  });

  // ==========================================
  // 1. EXPORT FUNCTIONALITY
  // ==========================================
  const copyBtn = document.getElementById('copyBtn');
  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      updateStatus("Reading template...", "info");
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab || !tab.url || !tab.url.toLowerCase().includes("listedit.aspx")) {
          updateStatus("Error: Must be on List Settings page!", "error");
          return;
        }

        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: scrapeSharePointColumns,
          world: 'MAIN'
        }, (results) => {
          if (!results || !results[0]) {
            updateStatus("Failed to run script on tab.", "error");
            return;
          }
          const data = results[0].result;
          if (data && data.error) {
            updateStatus(`Error: ${data.error}`, "error");
            return;
          }

          navigator.clipboard.writeText(JSON.stringify(data, null, 2))
            .then(() => updateStatus(`Copied ${data.length} columns to clipboard!`, "success"))
            .catch(err => updateStatus(`Clipboard failed: ${err.message}`, "error"));
        });
      } catch (err) {
        updateStatus(`Execution error: ${err.message}`, "error");
      }
    });
  }

  const removeColumnsBtn = document.getElementById('removeColumnsBtn');
  if (removeColumnsBtn) {
    removeColumnsBtn.addEventListener('click', async () => {
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab || !tab.url || !tab.url.toLowerCase().includes("listedit.aspx")) {
          updateStatus("Error: Must be on List Settings page!", "error");
          return;
        }

        const confirmed = window.confirm(
          "This will PERMANENTLY delete every custom column (and its data) from this list, keeping only built-in system columns. This cannot be undone. Continue?"
        );
        if (!confirmed) return;

        updateStatus("Removing non-system columns...", "info");

        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: removeNonSystemColumns,
          world: 'MAIN'
        }, (results) => {
          if (!results || !results[0]) {
            updateStatus("Failed to run script on tab.", "error");
            return;
          }
          const res = results[0].result;
          if (!res.success) {
            updateStatus(`Error: ${res.error}`, "error");
            return;
          }

          const failed = res.errors && res.errors.length ? ` ${res.errors.length} column(s) failed: ${res.errors.join(' | ')}` : "";
          updateStatus(`Removed ${res.deleted} column(s).${failed}`, failed ? "info" : "success");
        });
      } catch (err) {
        updateStatus(`Execution error: ${err.message}`, "error");
      }
    });
  }

  // ==========================================
  // 2. IMPORT FUNCTIONALITY
  // ==========================================
  const createBtn = document.getElementById('createBtn');
  if (createBtn) {
    createBtn.addEventListener('click', async () => {
      const title = document.getElementById('listTitle').value.trim();
      const rawJson = document.getElementById('jsonInput').value.trim();

      if (!title) return updateStatus("Please enter a list title.", "error");
      if (!rawJson) return updateStatus("Please paste JSON data.", "error");

      let columnData;
      try {
        columnData = JSON.parse(rawJson);
      } catch (e) {
        return updateStatus(`Invalid JSON structure: ${e.message}`, "error");
      }

      updateStatus("Provisioning list on page...", "info");
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: provisionSharePointList,
        args: [title, columnData],
        world: 'MAIN'
      }, (results) => {
        if (!results || !results[0]) {
          updateStatus("No confirmation back from the page.", "error");
          return;
        }
        const res = results[0].result;
        if (res.success) {
          const skipped = res.errors && res.errors.length ? ` ${res.errors.length} column(s) skipped: ${res.errors.join(' | ')}` : "";
          updateStatus(`Created list with ${res.count} columns.${skipped}`, skipped ? "info" : "success");
        } else {
          updateStatus(`Error: ${res.error}`, "error");
        }
      });
    });
  }

  // ==========================================
  // 2b. ADD COLUMNS TO CURRENT SITE FUNCTIONALITY
  // ==========================================
  const addColumnsBtn = document.getElementById('addColumnsBtn');
  if (addColumnsBtn) {
    addColumnsBtn.addEventListener('click', async () => {
      const prefix = document.getElementById('addColumnsPrefix').value.trim();
      const group = document.getElementById('addColumnsGroup').value.trim();
      const rawJson = document.getElementById('addColumnsJsonInput').value.trim();

      if (!rawJson) return updateStatus("Please paste JSON data.", "error");

      let columnData;
      try {
        columnData = JSON.parse(rawJson);
      } catch (e) {
        return updateStatus(`Invalid JSON structure: ${e.message}`, "error");
      }

      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

        updateStatus("Adding columns to current site...", "info");

        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: addSharePointColumns,
          args: [columnData, prefix, group],
          world: 'MAIN'
        }, (results) => {
          if (!results || !results[0]) {
            updateStatus("No confirmation back from the page.", "error");
            return;
          }
          const res = results[0].result;
          if (res.success) {
            const skipped = res.errors && res.errors.length ? ` ${res.errors.length} column(s) skipped: ${res.errors.join(' | ')}` : "";
            updateStatus(`Added ${res.count} columns.${skipped}`, skipped ? "info" : "success");
          } else {
            updateStatus(`Error: ${res.error}`, "error");
          }
        });
      } catch (err) {
        updateStatus(`Execution error: ${err.message}`, "error");
      }
    });
  }

  // ==========================================
  // 3. COPY ITEM FUNCTIONALITY
  // ==========================================
  const copyItemBtn = document.getElementById('copyItemBtn');
  if (copyItemBtn) {
    copyItemBtn.addEventListener('click', async () => {
      updateStatus("Duplicating item...", "info");
      const manualId = document.getElementById('copyItemId').value.trim();

      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: duplicateSharePointItem,
          args: [manualId],
          world: 'MAIN'
        }, (results) => {
          if (!results || !results[0]) {
            updateStatus("Failed to communicate with page.", "error");
            return;
          }
          const res = results[0].result;
          if (res.error) {
            updateStatus(`Error: ${res.error}`, "error");
            return;
          }

          updateStatus(`Successfully duplicated item! New Item ID: ${res.newItemId}`, "success");
        });
      } catch (err) {
        updateStatus(`Execution error: ${err.message}`, "error");
      }
    });
  }

  // ==========================================
  // 4. VERSIONS FUNCTIONALITY
  // ==========================================
  const viewVersionsBtn = document.getElementById('viewVersionsBtn');
  if (viewVersionsBtn) {
    viewVersionsBtn.addEventListener('click', async () => {
      updateStatus("Fetching version history...", "info");
      const manualId = document.getElementById('itemId').value.trim();

      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: fetchSharePointVersions,
          args: [manualId],
          world: 'MAIN'
        }, (results) => {
          if (!results || !results[0]) {
            updateStatus("Failed to communicate with page.", "error");
            return;
          }
          const res = results[0].result;
          if (res.error) {
            updateStatus(`Error: ${res.error}`, "error");
            return;
          }

          renderVersionTable(res.versions);
          updateStatus(`Successfully loaded ${res.versions.length} versions.`, "success");
        });
      } catch (err) {
        updateStatus(`Execution error: ${err.message}`, "error");
      }
    });
  }

  // ==========================================
  // 5. CONTENT TYPE MANAGER FUNCTIONALITY
  // ==========================================
  const fetchCtBtn = document.getElementById('fetchCtBtn');
  const ctFilterInput = document.getElementById('ctFilterInput');
  const ctFilterContainer = document.getElementById('ctFilterContainer');

  if (fetchCtBtn) {
    fetchCtBtn.addEventListener('click', async () => {
      updateStatus("Querying content types...", "info");
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: fetchSharePointContentTypes,
          world: 'MAIN'
        }, (results) => {
          if (!results || !results[0]) {
            updateStatus("Failed to fetch Content Types.", "error");
            return;
          }
          const res = results[0].result;
          if (res.error) {
            updateStatus(`Error: ${res.error}`, "error");
            return;
          }

          window.cachedContentTypes = res.contentTypes;
          if (ctFilterInput) ctFilterInput.value = ""; // Reset filter text
          if (ctFilterContainer) ctFilterContainer.style.display = 'block'; // Show filter UI

          renderContentTypeTable(res.contentTypes);
          updateStatus(`Loaded ${res.contentTypes.length} content types sorted by title. Click a row to see fields.`, "success");
        });
      } catch (err) {
        updateStatus(`Execution error: ${err.message}`, "error");
      }
    });
  }

  if (ctFilterInput) {
    ctFilterInput.addEventListener('input', () => {
      if (!window.cachedContentTypes) return;
      const query = ctFilterInput.value.toLowerCase();
      const filtered = window.cachedContentTypes.filter(ct =>
        ct.title.toLowerCase().includes(query)
      );
      renderContentTypeTable(filtered);
    });
  }
}

// Safely execute whether DOM is already ready or loading
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
