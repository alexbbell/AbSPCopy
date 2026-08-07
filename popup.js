document.addEventListener('DOMContentLoaded', () => {
  // --- TAB NAVIGATION SWITCHING ---
  const exportTabBtn = document.getElementById('tab-export');
  const importTabBtn = document.getElementById('tab-import');
  const copyItemTabBtn = document.getElementById('tab-copy-item');
  const versionsTabBtn = document.getElementById('tab-versions');
  const ctManagerTabBtn = document.getElementById('tab-ct-manager');

  const exportContent = document.getElementById('content-export');
  const importContent = document.getElementById('content-import');
  const copyItemContent = document.getElementById('content-copy-item');
  const versionsContent = document.getElementById('content-versions');
  const ctManagerContent = document.getElementById('content-ct-manager');

  function switchTab(activeBtn, activeContent) {
    [exportTabBtn, importTabBtn, copyItemTabBtn, versionsTabBtn, ctManagerTabBtn].forEach(btn => {
      if (btn) btn.classList.remove('active');
    });
    [exportContent, importContent, copyItemContent, versionsContent, ctManagerContent].forEach(content => {
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
  if (copyItemTabBtn) copyItemTabBtn.addEventListener('click', () => switchTab(copyItemTabBtn, copyItemContent));
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

  const rows = container.querySelectorAll('.ct-row');
  rows.forEach((row, idx) => {
    row.addEventListener('click', () => {
      rows.forEach(r => r.style.backgroundColor = 'white');
      row.style.backgroundColor = '#f3f2f1';
      renderContentTypeFields(contentTypes[idx]);
    });
  });
}

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

// ==========================================
// INJECTED FUNCTION: DUPLICATE LIST ITEM
// ==========================================
async function duplicateSharePointItem(manualId) {
  try {
    const urlParams = new URLSearchParams(window.location.search);
    let listId = urlParams.get('List');
    let itemId = manualId || urlParams.get('ID');

    if (!listId && typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo.pageListId) {
      listId = _spPageContextInfo.pageListId;
    }
    if (!listId && window.Ctx) {
      listId = window.Ctx.listName;
    }

    if (!listId) throw new Error("Could not detect List ID from context/URL.");
    if (!itemId) throw new Error("Could not detect Item ID from URL. Please enter it manually.");

    listId = listId.replace(/[{}]/g, "");
    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");

    // 1. Get Request Digest Token
    let formDigest = _spPageContextInfo.formDigestValue || (document.getElementById('__REQUESTDIGEST') ? document.getElementById('__REQUESTDIGEST').value : null);
    if (!formDigest) {
      const digestRes = await fetch(`${siteUrl}/_api/contextinfo`, {
        method: "POST",
        headers: { "Accept": "application/json;odata=nometadata" }
      });
      if (digestRes.ok) {
        const digestData = await digestRes.json();
        formDigest = digestData.FormDigestValue;
      }
    }
    if (!formDigest) throw new Error("Unable to obtain form digest token for operation authorization.");

    // 2. Retrieve non-readonly, non-hidden custom fields definition
    const fieldsRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!fieldsRes.ok) throw new Error("Failed to fetch writable list fields.");
    const fieldsData = await fieldsRes.json();

    const excludedSystemFields = [
      "contenttype", "attachments", "folderchildcount", "itemchildcount", 
      "edit", "linktitlenomenu", "linktitle", "docicon", "appauthor", "appeditor"
    ];

    const writableFields = fieldsData.value.filter(f => !excludedSystemFields.includes(f.InternalName.toLowerCase()));

    // 3. Fetch source item details
    const itemRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/items(${itemId})`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!itemRes.ok) throw new Error(`Could not locate source item ID #${itemId}.`);
    const sourceItem = await itemRes.json();

    // 4. Construct payload for new cloned item
    const newItemPayload = {};
    writableFields.forEach(field => {
      const name = field.InternalName;
      if (sourceItem.hasOwnProperty(name) && sourceItem[name] !== null && sourceItem[name] !== undefined) {
        newItemPayload[name] = sourceItem[name];
      }
    });

    // 5. Post to creation REST endpoint
    const createRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/items`, {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": formDigest
      },
      body: JSON.stringify(newItemPayload)
    });

    if (!createRes.ok) {
      const errText = await createRes.text();
      throw new Error(`Item duplicate request failed: ${errText}`);
    }

    const createdItem = await createRes.json();
    return { error: null, newItemId: createdItem.Id || createdItem.ID };

  } catch (e) {
    return { error: e.message, newItemId: null };
  }
}



async function fetchSharePointVersions(manualId) {
  try {
    const urlParams = new URLSearchParams(window.location.search);
    let listId = urlParams.get('List');
    let itemId = manualId || urlParams.get('ID');

    if (!listId && typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo.pageListId) {
      listId = _spPageContextInfo.pageListId;
    }
    if (!listId && window.Ctx) {
      listId = window.Ctx.listName; 
    }
    if (!listId && window.WPQ2FormCtx && window.WPQ2FormCtx.ListAttributes) {
      listId = window.WPQ2FormCtx.ListAttributes.Id;
    }

    if (!listId) throw new Error("Could not detect List GUID. Try navigating to a single item page.");
    if (!itemId) throw new Error("Could not detect Item ID from URL. Please enter it manually.");

    listId = listId.replace(/[{}]/g, "");
    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");

    // 1. Fetch Field Display Name Map
    const fieldsRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/fields`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!fieldsRes.ok) throw new Error("Failed to query list fields definition.");
    const fieldsData = await fieldsRes.json();
    
    const userFieldMap = {};
    fieldsData.value.forEach(field => {
      userFieldMap[field.InternalName.toLowerCase()] = field.Title;
    });

    // 2. Query the full version timeline array via REST
    const restVersionsUrl = `${siteUrl}/_api/web/lists/getbyid('${listId}')/items(${itemId})/versions`;
    const res = await fetch(restVersionsUrl, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!res.ok) throw new Error(`REST version stream rejected with status ${res.status}`);
    const data = await res.json();

    if (!data.value || data.value.length === 0) {
      return { error: null, versions: [] };
    }

    // 3. Format and clean JSON properties
    const formattedVersions = data.value.map(v => {
      const fieldsChanged = {};
      
      Object.keys(v).forEach(key => {
        const lowerKey = key.toLowerCase();
        
        // --- STRICT FILTERING FOR ODATA & SYSTEM FIELDS ---
        if (lowerKey.includes('odata') || lowerKey.startsWith('__')) return;
        
        // Skip system structural metadata fields
        const systemExclusions = [
          "iscurrentversion", "versionid", "versionlabel", "editor", "modified", "created",
          "id", "author", "attachments", "guid", "filesystemobjecttype", "serverredirectedurl"
        ];
        if (systemExclusions.includes(lowerKey)) return;
        
        // Skip link structures, metadata pointers, or operational hashes
        if (lowerKey.endsWith('meta') || lowerKey.endsWith('stringid') || lowerKey.endsWith('hash')) return;

        let val = v[key];
        if (val === null || val === undefined || val === "") return;

        let cleanVal = typeof val === 'object' ? JSON.stringify(val) : String(val);

        // Skip fields that just output empty/generic arrays or objects
        if (cleanVal === '[]' || cleanVal === '{}') return;

        // Pretty print JSON structures (such as multi-choice comment grids or workflow logs)
        if (cleanVal.startsWith('[') || cleanVal.startsWith('{')) {
          try {
            const parsed = JSON.parse(cleanVal);
            if (Array.isArray(parsed)) {
              cleanVal = parsed.map(item => item.text || JSON.stringify(item)).join("<br/>");
            }
          } catch(e) {}
        }

        // Translate the internal name into the clean display title
        const displayLabel = userFieldMap[lowerKey] || key;
        fieldsChanged[displayLabel] = cleanVal;
      });

      // Handle the version modifier info cleanly
      let editorUser = "System Account";
      if (v.Editor) {
        editorUser = v.Editor.LookupValue || v.Editor.Email || (typeof v.Editor === 'object' ? v.Editor.Title : String(v.Editor));
      }

      return {
        versionLabel: v.VersionLabel || "?.0",
        editor: editorUser,
        modified: v.Modified ? new Date(v.Modified).toLocaleString() : "",
        fields: fieldsChanged
      };
    });

    formattedVersions.sort((a, b) => parseFloat(b.versionLabel) - parseFloat(a.versionLabel));
    return { error: null, versions: formattedVersions };

  } catch (e) {
    return { error: e.message, versions: [] };
  }
}