document.addEventListener('DOMContentLoaded', () => {
  // --- TAB NAVIGATION SWITCHING ---
  const exportTabBtn = document.getElementById('tab-export');
  const importTabBtn = document.getElementById('tab-import');
  const versionsTabBtn = document.getElementById('tab-versions');

  const exportContent = document.getElementById('content-export');
  const importContent = document.getElementById('content-import');
  const versionsContent = document.getElementById('content-versions');

  function switchTab(activeBtn, activeContent) {
    [exportTabBtn, importTabBtn, versionsTabBtn].forEach(btn => {
      if (btn) btn.classList.remove('active');
    });
    [exportContent, importContent, versionsContent].forEach(content => {
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

// ==========================================
// PAGE CONTEXT SCRIPTS (Runs inside SharePoint)
// ==========================================
async function scrapeSharePointColumns() {
  try {
    const urlParams = new URLSearchParams(window.location.search);
    let listId = urlParams.get('List');
    if (!listId) throw new Error("Could not parse List GUID from page URL.");
    listId = listId.replace(/[{}]/g, "");

    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");
    const res = await fetch(`${siteUrl}/_api/web/lists/getbyid('${encodeURIComponent(listId)}')/fields`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!res.ok) throw new Error(`API returned HTTP ${res.status}`);
    const data = await res.json();

    return data.value.map(field => {
      if (field.FromBaseType || field.Hidden || field.ReadOnlyField) return null;
      if (["ID", "Title", "Created", "Modified", "Author", "Editor", "Attachments"].includes(field.InternalName)) return null;

      const info = { displayName: field.Title, internalName: field.InternalName, type: field.TypeAsString };
      if (field.TypeAsString === "Lookup" || field.TypeAsString === "LookupMulti") {
        info.lookupList = field.LookupList;
        info.lookupField = field.LookupField;
      }
      return info;
    }).filter(Boolean);
  } catch (e) {
    return { error: e.message };
  }
}

async function provisionSharePointList(listTitle, columns) {
  try {
    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");
    const contextRes = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!contextRes.ok) throw new Error("Could not acquire FormDigest token.");
    const contextData = await contextRes.json();
    const formDigest = contextData.FormDigestValue;

    const listResponse = await fetch(`${siteUrl}/_api/web/lists`, {
      method: "POST",
      body: JSON.stringify({ 'AllowContentTypes': true, 'BaseTemplate': 100, 'Title': listTitle }),
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": formDigest
      }
    });
    if (!listResponse.ok) {
      const txt = await listResponse.text();
      let m = txt; try { m = JSON.parse(txt)['odata.error'].message.value; } catch(e){}
      throw new Error(`List rejection: ${m}`);
    }
    const listData = await listResponse.json();
    const listId = listData.Id;

    const createFieldUrl = `${siteUrl}/_api/web/lists/getbyid('${listId}')/fields/CreateFieldAsXml`;
    let count = 0;
    for (const col of columns) {
      let xml = "";
      if (col.type === "Lookup" || col.type === "LookupMulti") {
        xml = `<Field Type="${col.type}" Name="${col.internalName}" DisplayName="${col.displayName}" List="${col.lookupList}" ShowField="${col.lookupField}" Mult="${col.type === "LookupMulti" ? "TRUE" : "FALSE"}"/>`;
      } else {
        xml = `<Field Type="${col.type}" Name="${col.internalName}" DisplayName="${col.displayName}"/>`;
      }

      const fRes = await fetch(createFieldUrl, {
        method: "POST",
        body: JSON.stringify({ parameters: { SchemaXml: xml } }),
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": formDigest
        }
      });
      if (!fRes.ok) {
        const fTxt = await fRes.text();
        let fM = fTxt; try { fM = JSON.parse(fTxt)['odata.error'].message.value; } catch(e){}
        throw new Error(`Column ${col.displayName} failed: ${fM}`);
      }
      count++;
    }
    return { success: true, count: count };
  } catch (err) {
    return { success: false, error: err.message };
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