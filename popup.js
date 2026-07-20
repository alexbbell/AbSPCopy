document.addEventListener('DOMContentLoaded', () => {
  // --- TAB NAVIGATION SWITCHING ---
  const exportTabBtn = document.getElementById('tab-export');
  const importTabBtn = document.getElementById('tab-import');
  const versionsTabBtn = document.getElementById('tab-versions');
  const ctManagerTabBtn = document.getElementById('tab-ct-manager');

  const exportContent = document.getElementById('content-export');
  const importContent = document.getElementById('content-import');
  const versionsContent = document.getElementById('content-versions');
  const ctManagerContent = document.getElementById('content-ct-manager');

  function switchTab(activeBtn, activeContent) {
    [exportTabBtn, importTabBtn, versionsTabBtn, ctManagerTabBtn].forEach(btn => {
      if (btn) btn.classList.remove('active');
    });
    [exportContent, importContent, versionsContent, ctManagerContent].forEach(content => {
      if (content) {
        content.classList.remove('active');
        content.style.display = 'none';
      }
    });
    if (activeBtn) activeBtn.classList.add('active');
    if (activeContent) {
      activeContent.classList.add('active');
      activeContent.style.display = 'block';
    }
  }

  if (exportTabBtn) exportTabBtn.addEventListener('click', () => switchTab(exportTabBtn, exportContent));
  if (importTabBtn) importTabBtn.addEventListener('click', () => switchTab(importTabBtn, importContent));
  if (versionsTabBtn) versionsTabBtn.addEventListener('click', () => switchTab(versionsTabBtn, versionsContent));
  if (ctManagerTabBtn) ctManagerTabBtn.addEventListener('click', () => switchTab(ctManagerTabBtn, ctManagerContent));

  // ==========================================
  // 1. EXPORT FUNCTIONALITY
  // ==========================================
  const copyBtn = document.getElementById('copyBtn');
  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      updateStatus("Reading template...", "info");
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab || !tab.url || !tab.url.includes("listedit.aspx")) {
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

  // ==========================================
// BACKGROUND INJECTED SCRIPTS (MAIN WORLD)
// ==========================================

async function scrapeSharePointColumns() {
  try {
    // 1. Ensure _spPageContextInfo is present
    if (typeof _spPageContextInfo === 'undefined') {
      throw new Error("SharePoint context (_spPageContextInfo) not found.");
    }

    // 2. Extract the List ID from the URL query parameters (common on listedit.aspx)
    const urlParams = new URLSearchParams(window.location.search);
    let listId = urlParams.get('List') || urlParams.get('list');
    
    // Fallback to page context if URL doesn't yield it
    if (!listId && _spPageContextInfo.pageListId) {
      listId = _spPageContextInfo.pageListId;
    }

    if (!listId) {
      throw new Error("Could not detect the SharePoint List ID from this page.");
    }

    // Clean up braces if present in the GUID
    listId = listId.replace(/[{}]/g, "");

    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");
    
    // 3. Query the fields, filtering out built-in hidden/readonly system fields 
    // We explicitly request SchemaXml for CAML XML dynamic provisioning downstream
    const restUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=Title,InternalName,TypeAsString,SchemaXml,LookupList,LookupField`;

    const res = await fetch(restUrl, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });

    if (!res.ok) {
      throw new Error(`SharePoint API rejected schema query with status ${res.status}`);
    }

    const data = await res.json();
    
    // 4. Map the payload to the layout specifications expected by the provision pipeline
    return data.value.map(field => ({
      title: field.Title,
      internalName: field.InternalName,
      type: field.TypeAsString,
      schemaXml: field.SchemaXml,
      lookupList: field.LookupList || null,
      lookupField: field.LookupField || null
    }));

  } catch (e) {
    return { error: e.message };
  }
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
          updateStatus(`Success! Created list with ${res.count} columns.`, "success");
        } else {
          updateStatus(`Error: ${res.error}`, "error");
        }
      });
    });
  }

  // ==========================================
  // 3. VERSIONS FUNCTIONALITY
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
  // 4. CONTENT TYPE MANAGER FUNCTIONALITY
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

  // Live filtering input watcher
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
});

function updateStatus(msg, type) {
  const statusEl = document.getElementById('status');
  if (!statusEl) return;
  statusEl.innerText = msg;
  statusEl.className = 'status ' + type;
}

function renderVersionTable(versions) {
  const container = document.getElementById('versionTableContainer');
  if (!container) return;
  if (versions.length === 0) {
    container.innerHTML = "<p style='font-size:12px;'>No versions found.</p>";
    return;
  }

  let html = `<table style="width:100%; border-collapse: collapse; font-size: 11px; background: white;">
    <thead>
      <tr style="background: #edebe9; text-align: left;">
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 40px;">Ver.</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 120px;">Modified By</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 130px;">Date</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce;">Changes / Content</th>
      </tr>
    </thead>
    <tbody>`;

  versions.forEach(v => {
    let changesHtml = "";
    for (const [key, val] of Object.entries(v.fields)) {
      changesHtml += `<strong>${key}:</strong> ${val}<br/>`;
    }

    html += `<tr style="border-bottom: 1px solid #edebe9;">
      <td style="padding: 6px; border: 1px solid #d2d0ce; vertical-align: top;">${v.versionLabel}</td>
      <td style="padding: 6px; border: 1px solid #d2d0ce; vertical-align: top;">${v.editor}</td>
      <td style="padding: 6px; border: 1px solid #d2d0ce; vertical-align: top; white-space: nowrap;">${v.modified}</td>
      <td style="padding: 6px; border: 1px solid #d2d0ce; vertical-align: top;">${changesHtml || '<em>No custom fields changed</em>'}</td>
    </tr>`;
  });

  html += `</tbody></table>`;
  container.innerHTML = html;
}

// Render Content Type master table
function renderContentTypeTable(contentTypes) {
  const container = document.getElementById('ctTableContainer');
  const fieldsContainer = document.getElementById('ctFieldsContainer');
  if (!container) return;
  
  fieldsContainer.innerHTML = ""; 

  if (contentTypes.length === 0) {
    container.innerHTML = "<p style='font-size:12px; padding:10px;'>No matching content types found.</p>";
    return;
  }

  let html = `<table style="width:100%; border-collapse: collapse; font-size: 11px; background: white; cursor: pointer;">
    <thead>
      <tr style="background: #edebe9; text-align: left; position: sticky; top: 0; box-shadow: inset 0 -1px 0 #d2d0ce;">
        <th style="padding: 6px; border: 1px solid #d2d0ce;">CTID</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce;">Title</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce;">Group</th>
      </tr>
    </thead>
    <tbody>`;

  contentTypes.forEach((ct, index) => {
    html += `<tr class="ct-row" data-index="${index}" style="border-bottom: 1px solid #edebe9;">
      <td style="padding: 6px; border: 1px solid #d2d0ce; font-family: Courier, monospace; font-size:10px; word-break: break-all;">${ct.id}</td>
      <td style="padding: 6px; border: 1px solid #d2d0ce; font-weight: 600; color: #0078d4;">${ct.title}</td>
      <td style="padding: 6px; border: 1px solid #d2d0ce; color: #484644;">${ct.group}</td>
    </tr>`;
  });

  html += `</tbody></table>`;
  container.innerHTML = html;

  // Row event processing maps directly to the active array context
  const rows = container.querySelectorAll('.ct-row');
  rows.forEach((row, idx) => {
    row.addEventListener('click', () => {
      rows.forEach(r => r.style.backgroundColor = 'white');
      row.style.backgroundColor = '#f3f2f1';
      renderContentTypeFields(contentTypes[idx]);
    });
  });
}



// Render fields array for a specific selected Content Type
function renderContentTypeFields(ct) {
  const container = document.getElementById('ctFieldsContainer');
  if (!container) return;

  let html = `<h4 style="margin: 0 0 8px 0; font-size: 12px; color: #323130;">Fields for: <span style="color:#0078d4;">${ct.title}</span></h4>
    <table style="width:100%; border-collapse: collapse; font-size: 11px; background: white; border: 1px solid #d2d0ce;">
    <thead>
      <tr style="background: #f3f2f1; text-align: left;">
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 30%;">Title</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 25%;">Internal Name</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 20%;">Type</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 25%;">Field ID</th>
      </tr>
    </thead>
    <tbody>`;

  if (!ct.fields || ct.fields.length === 0) {
    html += `<tr><td colspan="4" style="padding: 8px; text-align: center; color: #a19f9d; font-style: italic;">No fields mapped to this content type structure.</td></tr>`;
  } else {
    ct.fields.forEach(f => {
      html += `<tr style="border-bottom: 1px solid #edebe9;">
        <td style="padding: 6px; border: 1px solid #d2d0ce; font-weight: 600; color: #323130;">${f.title}</td>
        <td style="padding: 6px; border: 1px solid #d2d0ce; font-family: Courier, monospace; font-size:10px; color: #0078d4;">${f.internalName}</td>
        <td style="padding: 6px; border: 1px solid #d2d0ce; color: #484644;">${f.type}</td>
        <td style="padding: 6px; border: 1px solid #d2d0ce; font-family: Courier, monospace; font-size:10px; color: #605e5c;">${f.id}</td>
      </tr>`;
    });
  }

  html += `</tbody></table>`;
  container.innerHTML = html;
}

// REST Pipeline target querying Content Types expanded with system field relationships
async function fetchSharePointContentTypes() {
  try {
    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");
    // Explicitly select and expand fields to include TypeAsString and InternalName properties
    const restUrl = `${siteUrl}/_api/web/contenttypes?$select=StringId,Name,Group,Fields/Id,Fields/Title,Fields/InternalName,Fields/TypeAsString&$expand=Fields`;
    
    const res = await fetch(restUrl, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!res.ok) throw new Error(`Content type API rejected with status code ${res.status}`);
    const data = await res.json();

    const formattedCts = data.value.map(ct => {
      return {
        id: ct.StringId || (ct.Id ? ct.Id.StringId : ""),
        title: ct.Name,
        group: ct.Group,
        fields: (ct.Fields || []).map(f => ({
          id: f.Id,
          title: f.Title,
          internalName: f.InternalName || "",
          type: f.TypeAsString || "Unknown"
        }))
      };
    });

    // Alphabetical sort by Title field
    formattedCts.sort((a, b) => a.title.localeCompare(b.title));

    return { error: null, contentTypes: formattedCts };``
  } catch (e) {
    return { error: e.message, contentTypes: [] };
  }
}