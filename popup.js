// --- TAB NAVIGATION SWITCHING ---
const exportTabBtn = document.getElementById('tab-export');
const importTabBtn = document.getElementById('tab-import');
const exportContent = document.getElementById('content-export');
const importContent = document.getElementById('content-import');

exportTabBtn.addEventListener('click', () => {
  exportTabBtn.classList.add('active');
  importTabBtn.classList.remove('active');
  exportContent.classList.add('active');
  importContent.classList.remove('active');
});

importTabBtn.addEventListener('click', () => {
  importTabBtn.classList.add('active');
  exportTabBtn.classList.remove('active');
  importContent.classList.add('active');
  exportContent.classList.remove('active');
});

// Helper to update the extension popup UI status text
function updateStatus(msg, type) {
  const statusEl = document.getElementById('status');
  if (!statusEl) return;
  statusEl.innerText = msg;
  statusEl.className = 'status ' + type;
}

// ==========================================
// 1. EXPORT FUNCTIONALITY (Copy Columns)
// ==========================================
document.getElementById('copyBtn').addEventListener('click', async () => {
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

// ==========================================
// 2. IMPORT FUNCTIONALITY (Create List)
// ==========================================
document.getElementById('createBtn').addEventListener('click', async () => {
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