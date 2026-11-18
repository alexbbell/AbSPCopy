// injectedScripts.js

// NOTE: this function is passed directly to chrome.scripting.executeScript({ func: ... }),
// which serializes only its own source and runs it standalone in the target page's world.
// It must not reference anything declared outside itself (no outer consts/helpers).
export async function provisionSharePointList(listTitle, columns) {
  try {
    // Human-readable type labels (as shown on the classic List Settings page) mapped
    // to the internal SP field type names required by CreateFieldAsXml.
    const FIELD_TYPE_LABEL_MAP = {
      "single line of text": "Text",
      "multiple lines of text": "Note",
      "choice": "Choice",
      "number": "Number",
      "currency": "Currency",
      "date and time": "DateTime",
      "lookup": "Lookup",
      "yes/no": "Boolean",
      "person or group": "User",
      "hyperlink or picture": "URL",
      "calculated": "Calculated",
      "task outcome": "OutcomeChoice",
      "external data": "ExternalData",
      "managed metadata": "TaxonomyFieldType"
    };

    // Fields present by default on any generic list, or types that CreateFieldAsXml
    // cannot provision from a plain schema (computed formulas, attachments, etc.).
    const SKIPPED_FIELDS = new Set(["title", "contenttype", "attachments", "id", "created", "modified", "author", "editor"]);
    const SKIPPED_TYPES = new Set(["computed", "attachments", "contenttype"]);

    function resolveFieldType(col) {
      const raw = col.type || col.TypeAsString;
      if (raw) return raw;
      if (col.FieldTypeKind) {
        return FIELD_TYPE_LABEL_MAP[String(col.FieldTypeKind).trim().toLowerCase()] || null;
      }
      return null;
    }

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
    const errors = [];

    for (const col of columns) {
      const displayName = col.displayName || col.Title || col.title || "";
      const internalName = col.internalName || col.InternalName || displayName.replace(/[^A-Za-z0-9_]/g, "");
      const type = resolveFieldType(col);

      if (SKIPPED_FIELDS.has(internalName.toLowerCase()) || (type && SKIPPED_TYPES.has(type.toLowerCase()))) {
        continue; // built-in field already present on every new list, or not creatable from a plain schema
      }

      if (!type) {
        errors.push(`${displayName || internalName}: could not determine a field type to create.`);
        continue;
      }

      let xml = "";
      if (type === "Lookup" || type === "LookupMulti") {
        const lookupList = col.lookupList || col.LookupList;
        const lookupField = col.lookupField || col.LookupField;
        if (!lookupList || !lookupField) {
          errors.push(`${displayName}: Lookup column is missing its source list/field reference and must be added manually.`);
          continue;
        }
        xml = `<Field Type="${type}" Name="${internalName}" DisplayName="${displayName}" List="${lookupList}" ShowField="${lookupField}" Mult="${type === "LookupMulti" ? "TRUE" : "FALSE"}"/>`;
      } else {
        xml = `<Field Type="${type}" Name="${internalName}" DisplayName="${displayName}"/>`;
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
        errors.push(`${displayName}: ${fM}`);
        continue;
      }
      count++;
    }
    return { success: true, count: count, errors: errors };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// NOTE: this function is passed directly to chrome.scripting.executeScript({ func: ... }),
// which serializes only its own source and runs it standalone in the target page's world.
// It must not reference anything declared outside itself (no outer consts/helpers).
export async function addSharePointColumns(columns, prefix, group) {
  try {
    // Human-readable type labels (as shown on the classic List Settings page) mapped
    // to the internal SP field type names required by CreateFieldAsXml.
    const FIELD_TYPE_LABEL_MAP = {
      "single line of text": "Text",
      "multiple lines of text": "Note",
      "choice": "Choice",
      "number": "Number",
      "currency": "Currency",
      "date and time": "DateTime",
      "lookup": "Lookup",
      "yes/no": "Boolean",
      "person or group": "User",
      "hyperlink or picture": "URL",
      "calculated": "Calculated",
      "task outcome": "OutcomeChoice",
      "external data": "ExternalData",
      "managed metadata": "TaxonomyFieldType"
    };

    // Fields present by default on any generic list, or types that CreateFieldAsXml
    // cannot provision from a plain schema (computed formulas, attachments, etc.).
    const SKIPPED_FIELDS = new Set(["title", "contenttype", "attachments", "id", "created", "modified", "author", "editor"]);
    const SKIPPED_TYPES = new Set(["computed", "attachments", "contenttype"]);

    function resolveFieldType(col) {
      const raw = col.type || col.TypeAsString;
      if (raw) return raw;
      if (col.FieldTypeKind) {
        return FIELD_TYPE_LABEL_MAP[String(col.FieldTypeKind).trim().toLowerCase()] || null;
      }
      return null;
    }

    function escapeXmlAttr(str) {
      return String(str).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    if (typeof _spPageContextInfo === 'undefined') {
      throw new Error("Could not detect SharePoint site context.");
    }

    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");

    const contextRes = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!contextRes.ok) throw new Error("Could not acquire FormDigest token.");
    const contextData = await contextRes.json();
    const formDigest = contextData.FormDigestValue;

    // Site-scoped (not list-scoped) so the Name attribute is honored as the real internal
    // name and the fields show up as reusable Site Columns under the requested Group.
    const createFieldUrl = `${siteUrl}/_api/web/fields/CreateFieldAsXml`;
    const prefixStr = prefix ? String(prefix) : "";
    const groupStr = group ? String(group) : "";
    const groupAttr = groupStr ? ` Group="${escapeXmlAttr(groupStr)}"` : "";

    let count = 0;
    const errors = [];

    for (const col of columns) {
      const displayName = col.displayName || col.Title || col.title || "";
      const baseInternalName = col.internalName || col.InternalName || displayName.replace(/[^A-Za-z0-9_]/g, "");
      const internalName = prefixStr + baseInternalName;
      const type = resolveFieldType(col);

      if (SKIPPED_FIELDS.has(baseInternalName.toLowerCase()) || (type && SKIPPED_TYPES.has(type.toLowerCase()))) {
        continue; // built-in field already present on every list, or not creatable from a plain schema
      }

      if (!type) {
        errors.push(`${displayName || internalName}: could not determine a field type to create.`);
        continue;
      }

      const required = (col.Required || col.required) ? ' Required="TRUE"' : '';
      const safeDisplayName = escapeXmlAttr(displayName);

      let xml = "";
      if (type === "Lookup" || type === "LookupMulti") {
        const lookupList = col.lookupList || col.LookupList;
        const lookupField = col.lookupField || col.LookupField;
        if (!lookupList || !lookupField) {
          errors.push(`${displayName}: Lookup column is missing its source list/field reference and must be added manually.`);
          continue;
        }
        xml = `<Field Type="${type}" Name="${internalName}" DisplayName="${safeDisplayName}" List="${lookupList}" ShowField="${lookupField}" Mult="${type === "LookupMulti" ? "TRUE" : "FALSE"}"${groupAttr}${required}/>`;
      } else if (type === "User" || type === "UserMulti") {
        // Person/Group fields are internally a Lookup into the hidden UserInfo list.
        xml = `<Field Type="${type}" Name="${internalName}" DisplayName="${safeDisplayName}" List="UserInfo"${type === "UserMulti" ? ' Mult="TRUE"' : ''}${groupAttr}${required}/>`;
      } else {
        xml = `<Field Type="${type}" Name="${internalName}" DisplayName="${safeDisplayName}"${groupAttr}${required}/>`;
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
        let fM = fTxt; try { fM = JSON.parse(fTxt)['odata.error'].message.value; } catch (e) {}
        errors.push(`${displayName}: ${fM}`);
        continue;
      }

      // CreateFieldAsXml can fall back to the internal Name for the Title (e.g. when
      // DisplayName collides with an existing field elsewhere), so force the intended,
      // unprefixed Title explicitly as a follow-up update.
      const titleFixRes = await fetch(`${siteUrl}/_api/web/fields/getbyinternalnameortitle('${encodeURIComponent(internalName)}')`, {
        method: "POST",
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": formDigest,
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*"
        },
        body: JSON.stringify({ Title: displayName })
      });
      if (!titleFixRes.ok) {
        const tTxt = await titleFixRes.text();
        let tM = tTxt; try { tM = JSON.parse(tTxt)['odata.error'].message.value; } catch (e) {}
        errors.push(`${displayName}: created, but could not set Title: ${tM}`);
      }
      count++;
    }
    return { success: true, count: count, errors: errors };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Injectable: resolves the absolute URL of the SharePoint web the active page belongs to.
// Runs in the page (MAIN world); returns { webUrl } or { error }.
export async function resolveSharePointWebUrl() {
  try {
    if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo) {
      if (_spPageContextInfo.webAbsoluteUrl) return { webUrl: _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") };
      if (_spPageContextInfo.webServerRelativeUrl) {
        return { webUrl: window.location.origin + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "") };
      }
    }

    // Modern pages may not expose _spPageContextInfo: derive candidate web paths from the URL
    // and let SharePoint tell us the real web URL via /_api/contextinfo.
    const path = decodeURIComponent(window.location.pathname);
    const cutMarkers = ["/_layouts/", "/Lists/", "/SitePages/", "/Pages/", "/Forms/", "/Shared Documents", "/Freigegebene Dokumente"];
    let basePath = path;
    for (const marker of cutMarkers) {
      const i = basePath.toLowerCase().indexOf(marker.toLowerCase());
      if (i >= 0) basePath = basePath.substring(0, i);
    }
    const candidates = [basePath.replace(/\/[^\/]*\.aspx$/i, "").replace(/\/$/, "")];
    const m = path.match(/^\/(sites|teams|personal)\/[^\/]+/i);
    if (m) candidates.push(m[0]);
    candidates.push("");

    for (const c of [...new Set(candidates)]) {
      try {
        const r = await fetch(`${window.location.origin}${encodeURI(c)}/_api/contextinfo`, {
          method: "POST",
          headers: { "Accept": "application/json;odata=nometadata" }
        });
        if (!r.ok) continue;
        const d = await r.json();
        if (d && d.WebFullUrl) return { webUrl: d.WebFullUrl.replace(/\/$/, "") };
      } catch (e) { /* try next candidate */ }
    }
    return { error: "Could not detect SharePoint site context. Open a page of the SharePoint site and try again." };
  } catch (e) {
    return { error: e.message };
  }
}

export async function fetchSharePointContentTypes(webUrl) {
  try {
    if (!webUrl) throw new Error("SharePoint web URL was not provided.");
    const siteUrl = String(webUrl).replace(/\/$/, "");
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

    formattedCts.sort((a, b) => a.title.localeCompare(b.title));

    return { error: null, contentTypes: formattedCts };
  } catch (e) {
    return { error: e.message, contentTypes: [] };
  }
}

// Injectable: reads basic web info plus the full property bag (AllProperties) of the given web.
export async function fetchSharePointPropertyBag(webUrl) {
  try {
    if (!webUrl) throw new Error("SharePoint web URL was not provided.");
    const siteUrl = String(webUrl).replace(/\/$/, "");
    const headers = { "Accept": "application/json;odata=nometadata" };

    // REST encodes special characters in property names as _xHHHH_ (e.g. "_" -> "_x005f_").
    function decodeKey(key) {
      return key.replace(/_x([0-9a-fA-F]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
    }

    const [webRes, propsRes] = await Promise.all([
      fetch(`${siteUrl}/_api/web?$select=Id,Title,Url,ServerRelativeUrl,WebTemplate,Configuration,Language,Created,LastItemModifiedDate,UIVersion`, { headers }),
      fetch(`${siteUrl}/_api/web/AllProperties`, { headers })
    ]);

    if (!propsRes.ok) {
      const t = await propsRes.text();
      let msg = t; try { msg = JSON.parse(t)['odata.error'].message.value; } catch (e) {}
      throw new Error(`Property bag request failed (${propsRes.status}): ${msg}`);
    }
    const propsData = await propsRes.json();

    const properties = Object.keys(propsData)
      .filter(k => !k.startsWith("odata."))
      .map(k => {
        const v = propsData[k];
        return { key: decodeKey(k), value: v === null || v === undefined ? "" : (typeof v === "object" ? JSON.stringify(v) : String(v)) };
      })
      .sort((a, b) => a.key.localeCompare(b.key, undefined, { sensitivity: "base" }));

    let web = null;
    if (webRes.ok) {
      const w = await webRes.json();
      web = {
        title: w.Title,
        url: w.Url,
        serverRelativeUrl: w.ServerRelativeUrl,
        id: w.Id,
        template: `${w.WebTemplate}#${w.Configuration}`,
        language: w.Language,
        uiVersion: w.UIVersion,
        created: w.Created,
        lastItemModified: w.LastItemModifiedDate
      };
    }

    return { error: null, web: web, properties: properties };
  } catch (e) {
    return { error: e.message, web: null, properties: [] };
  }
}

export function scrapeSharePointColumns() {
  try {
    const columnRows = document.querySelectorAll("table[id*='rgList'] tr, table.ms-summarystandardbody tr");
    const columns = [];

    columnRows.forEach(row => {
      const link = row.querySelector("a[href*='FldEdit.aspx']");
      if (link) {
        const title = link.textContent.trim();
        const typeCell = row.querySelectorAll("td")[1];
        const type = typeCell ? typeCell.textContent.trim() : "Single line of text";

        columns.push({
          Title: title,
          FieldTypeKind: type
        });
      }
    });

    if (columns.length > 0) {
      return columns;
    }

    if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo.pageListId) {
      const siteUrl = _spPageContextInfo.webAbsoluteUrl;
      const listId = _spPageContextInfo.pageListId.replace(/[{}]/g, '');
      
      const xhr = new XMLHttpRequest();
      xhr.open("GET", `${siteUrl}/_api/web/lists/getbyid('${listId}')/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=Title,InternalName,TypeAsString,Required,LookupList,LookupField`, false);
      xhr.setRequestHeader("Accept", "application/json;odata=nometadata");
      xhr.send();

      if (xhr.status === 200) {
        const res = JSON.parse(xhr.responseText);
        return res.value.map(f => ({
          Title: f.Title,
          InternalName: f.InternalName,
          TypeAsString: f.TypeAsString,
          Required: f.Required,
          LookupList: f.LookupList,
          LookupField: f.LookupField
        }));
      }
    }

    return { error: "No columns found on page. Verify you are on List Settings." };
  } catch (e) {
    return { error: e.message };
  }
}

// NOTE: this function is passed directly to chrome.scripting.executeScript({ func: ... }),
// which serializes only its own source and runs it standalone in the target page's world.
// It must not reference anything declared outside itself (no outer consts/helpers).
export async function removeNonSystemColumns() {
  try {
    if (typeof _spPageContextInfo === 'undefined' || !_spPageContextInfo.pageListId) {
      throw new Error("Could not detect list. Navigate to the List Settings page first.");
    }

    const siteUrl = _spPageContextInfo.webAbsoluteUrl;
    const listId = _spPageContextInfo.pageListId.replace(/[{}]/g, '');

    const fieldsRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/fields?$select=Title,InternalName,Hidden,ReadOnlyField,FromBaseType,CanBeDeleted`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!fieldsRes.ok) throw new Error(`Failed to query list fields (status ${fieldsRes.status}).`);
    const fieldsData = await fieldsRes.json();

    // FromBaseType marks columns inherited from the list's base schema (Title, Created, Modified,
    // Author, Editor, Content Type, Attachments, etc.) - i.e. the "system" columns to keep.
    const candidates = fieldsData.value.filter(f => !f.Hidden && !f.ReadOnlyField && !f.FromBaseType && f.CanBeDeleted);

    if (candidates.length === 0) {
      return { success: true, deleted: 0, errors: [] };
    }

    const digestRes = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!digestRes.ok) throw new Error("Could not acquire FormDigest token.");
    const digestData = await digestRes.json();
    const formDigest = digestData.FormDigestValue;

    let deleted = 0;
    const errors = [];

    for (const field of candidates) {
      const delRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/fields/getbyinternalnameortitle('${encodeURIComponent(field.InternalName)}')`, {
        method: "POST",
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": formDigest,
          "X-HTTP-Method": "DELETE"
        }
      });
      if (!delRes.ok) {
        const txt = await delRes.text();
        let m = txt; try { m = JSON.parse(txt)['odata.error'].message.value; } catch (e) {}
        errors.push(`${field.Title || field.InternalName}: ${m}`);
        continue;
      }
      deleted++;
    }

    return { success: true, deleted: deleted, errors: errors };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

export async function duplicateSharePointItem(manualId) {
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

    const itemRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/items(${itemId})`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!itemRes.ok) throw new Error(`Could not locate source item ID #${itemId}.`);
    const sourceItem = await itemRes.json();

    const newItemPayload = {};
    writableFields.forEach(field => {
      const name = field.InternalName;
      if (sourceItem.hasOwnProperty(name) && sourceItem[name] !== null && sourceItem[name] !== undefined) {
        newItemPayload[name] = sourceItem[name];
      }
    });

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

// NOTE: this function is passed directly to chrome.scripting.executeScript({ func: ... }),
// which serializes only its own source and runs it standalone in the target page's world.
// It must not reference anything declared outside itself (no outer consts/helpers).
export async function exportSharePointItem(manualId) {
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

    const itemRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/items(${itemId})`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!itemRes.ok) throw new Error(`Could not locate source item ID #${itemId}.`);
    const sourceItem = await itemRes.json();

    const exportedItem = {};
    writableFields.forEach(field => {
      const name = field.InternalName;
      if (sourceItem.hasOwnProperty(name) && sourceItem[name] !== null && sourceItem[name] !== undefined) {
        exportedItem[name] = sourceItem[name];
      }
    });

    return { error: null, item: exportedItem, itemId: itemId };
  } catch (e) {
    return { error: e.message, item: null };
  }
}

// NOTE: this function is passed directly to chrome.scripting.executeScript({ func: ... }),
// which serializes only its own source and runs it standalone in the target page's world.
// It must not reference anything declared outside itself (no outer consts/helpers).
export async function importSharePointItem(itemData) {
  try {
    const urlParams = new URLSearchParams(window.location.search);
    let listId = urlParams.get('List');

    if (!listId && typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo.pageListId) {
      listId = _spPageContextInfo.pageListId;
    }
    if (!listId && window.Ctx) {
      listId = window.Ctx.listName;
    }

    if (!listId) throw new Error("Could not detect List ID. Open a page for the target list first.");

    listId = listId.replace(/[{}]/g, "");
    const siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");

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

    const fieldsRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`, {
      method: "GET",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!fieldsRes.ok) throw new Error("Failed to fetch writable list fields.");
    const fieldsData = await fieldsRes.json();
    const writableNames = new Set(fieldsData.value.map(f => f.InternalName));

    const payload = {};
    const skippedFields = [];
    Object.keys(itemData || {}).forEach(key => {
      if (writableNames.has(key)) {
        payload[key] = itemData[key];
      } else {
        skippedFields.push(key);
      }
    });

    const createRes = await fetch(`${siteUrl}/_api/web/lists/getbyid('${listId}')/items`, {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": formDigest
      },
      body: JSON.stringify(payload)
    });

    if (!createRes.ok) {
      const errText = await createRes.text();
      let m = errText; try { m = JSON.parse(errText)['odata.error'].message.value; } catch (e) {}
      throw new Error(m);
    }

    const createdItem = await createRes.json();
    return { error: null, newItemId: createdItem.Id || createdItem.ID, skippedFields: skippedFields };
  } catch (e) {
    return { error: e.message, newItemId: null };
  }
}

export async function fetchSharePointVersions(manualId) {
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

    const formattedVersions = data.value.map(v => {
      const fieldsChanged = {};
      
      Object.keys(v).forEach(key => {
        const lowerKey = key.toLowerCase();
        
        if (lowerKey.includes('odata') || lowerKey.startsWith('__')) return;
        
        const systemExclusions = [
          "iscurrentversion", "versionid", "versionlabel", "editor", "modified", "created",
          "id", "author", "attachments", "guid", "filesystemobjecttype", "serverredirectedurl"
        ];
        if (systemExclusions.includes(lowerKey)) return;
        if (lowerKey.endsWith('meta') || lowerKey.endsWith('stringid') || lowerKey.endsWith('hash')) return;

        let val = v[key];
        if (val === null || val === undefined || val === "") return;

        let cleanVal = typeof val === 'object' ? JSON.stringify(val) : String(val);

        if (cleanVal === '[]' || cleanVal === '{}') return;

        if (cleanVal.startsWith('[') || cleanVal.startsWith('{')) {
          try {
            const parsed = JSON.parse(cleanVal);
            if (Array.isArray(parsed)) {
              cleanVal = parsed.map(item => item.text || JSON.stringify(item)).join("<br/>");
            }
          } catch(e) {}
        }

        const displayLabel = userFieldMap[lowerKey] || key;
        fieldsChanged[displayLabel] = cleanVal;
      });

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