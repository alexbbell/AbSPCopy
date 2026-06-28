# SharePoint List & Column Cloner Extension

A lightweight, developer-focused Google Chrome Extension designed to seamlessly replicate SharePoint custom list schemas and document library column structures across lists or site collections. 

By leveraging the **SharePoint REST API** directly within the active browser tab session context, this tool extracts deep architectural layout specifications (including complex field variants such as lookups and multiline notes) and generates real-time provisioning commands without requiring any external OAuth authentication setups.

---

## 🚀 Key Architectural Features

- **Contextual Session Inheritance:** Executes web requests utilizing `world: 'MAIN'` execution states to capture live page parameters (`_spPageContextInfo`) and reuse implicit browser cookies/security tokens (`FormDigestValue`).
- **OData Non-Metadata Payloads:** Optimizes API transaction layers to communicate using `odata=nometadata` headers to bypass case-sensitive OData class wraps and maximize system performance.
- **Relational Lookup Tracking:** Extends past simple DOM-table scraping by accessing raw system tables to map hidden structural dependencies (`LookupList` targeting GUIDs and `LookupField` display variables).
- **CAML XML Schema Provisioning:** Bypasses basic `SP.Field` integer maps by utilizing server-side dynamic compilation through the native `/fields/CreateFieldAsXml` endpoint.

---

## 🛠️ Technical Implementation Blueprint

The extension follows a standard **Manifest V3** pattern, decoupling the frontend tab manipulation user interface from native page memory variables.


## Critical Field Schema Remapping Logic:
When processing incoming array definitions, standard REST schemas present items under human-readable designations (Note, Choice) that fail standard XML compiler assessments. The provisioning module normalizes these formats automatically inside the execution pipeline:

Note Fields: Maintained as safe string allocations or enriched with multiline layout definitions.

Choice Fields: Wrapped dynamically in safe baseline array elements (<CHOICES><CHOICE>Default Option</CHOICE></CHOICES>) to avoid schema compiler crashes caused by unpopulated target datasets.


## Author
Aleksei Beliaev
https://beliaeff.de
https://realdevtips.com

