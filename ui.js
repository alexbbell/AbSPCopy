// ui.js
export function updateStatus(msg, type) {
  const statusEl = document.getElementById('status');
  if (!statusEl) return;
  statusEl.innerText = msg;
  statusEl.className = 'status ' + type;
}

export function switchTab(activeBtn, activeContent, allBtns, allContents) {
  allBtns.forEach(btn => btn && btn.classList.remove('active'));
  allContents.forEach(content => {
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

export function renderVersionTable(versions) {
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

export function renderContentTypeTable(contentTypes) {
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

export function renderContentTypeFields(ct) {
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

function escapeHtml(str) {
  return String(str === null || str === undefined ? "" : str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

export function renderSiteInfo(web) {
  const container = document.getElementById('siteInfoContainer');
  if (!container) return;
  if (!web) {
    container.innerHTML = "";
    return;
  }

  const rows = [
    ["Title", web.title],
    ["URL", web.url],
    ["Server Relative URL", web.serverRelativeUrl],
    ["Web ID", web.id],
    ["Template", web.template],
    ["Language (LCID)", web.language],
    ["UI Version", web.uiVersion],
    ["Created", web.created ? new Date(web.created).toLocaleString() : ""],
    ["Last Item Modified", web.lastItemModified ? new Date(web.lastItemModified).toLocaleString() : ""]
  ];

  let html = `<table style="width:100%; border-collapse: collapse; font-size: 11px; background: white; border: 1px solid #d2d0ce;"><tbody>`;
  rows.forEach(([label, value]) => {
    html += `<tr>
      <td style="padding: 5px 6px; border: 1px solid #d2d0ce; background: #f3f2f1; font-weight: 600; width: 30%;">${escapeHtml(label)}</td>
      <td style="padding: 5px 6px; border: 1px solid #d2d0ce; word-break: break-all;">${escapeHtml(value)}</td>
    </tr>`;
  });
  html += `</tbody></table>`;
  container.innerHTML = html;
}

export function renderPropertyBagTable(properties) {
  const container = document.getElementById('propsTableContainer');
  if (!container) return;

  if (properties.length === 0) {
    container.innerHTML = "<p style='font-size:12px; padding:10px;'>No matching properties found.</p>";
    return;
  }

  let html = `<table style="width:100%; border-collapse: collapse; font-size: 11px; background: white; table-layout: fixed;">
    <thead>
      <tr style="background: #edebe9; text-align: left; position: sticky; top: 0; box-shadow: inset 0 -1px 0 #d2d0ce;">
        <th style="padding: 6px; border: 1px solid #d2d0ce; width: 38%;">Key</th>
        <th style="padding: 6px; border: 1px solid #d2d0ce;">Value</th>
      </tr>
    </thead>
    <tbody>`;

  properties.forEach(p => {
    html += `<tr style="border-bottom: 1px solid #edebe9;">
      <td style="padding: 6px; border: 1px solid #d2d0ce; font-family: Courier, monospace; font-size: 10px; color: #0078d4; word-break: break-all; vertical-align: top;">${escapeHtml(p.key)}</td>
      <td style="padding: 6px; border: 1px solid #d2d0ce; font-family: Courier, monospace; font-size: 10px; color: #323130; word-break: break-all; white-space: pre-wrap; vertical-align: top;">${p.value === "" ? '<em style="color:#a19f9d;">(empty)</em>' : escapeHtml(p.value)}</td>
    </tr>`;
  });

  html += `</tbody></table>`;
  container.innerHTML = html;
}
