# z/OS Compatibility Intelligence

A seven-page Streamlit dashboard built on the **original project** with its six vendor compatibility CSVs, existing matching/Excel utilities, and optional watsonx.ai RAG assistant. Includes a new read-only Broadcom HTML crawler, linked maintenance evidence extraction, versioned SQLite snapshots, and change detection.

## Start locally

Requires Python **3.11–3.12**, outbound HTTPS access to authorized vendor websites, and normal filesystem write permissions.

```bash
python -m venv .venv
source .venv/bin/activate           # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

The dashboard opens at `http://localhost:8501`. It immediately displays the **240 original vendor records**; no network scraping is required for first use.

## Seven working pages

1. **Overview** — counts, support status, vendor coverage, latest snapshot, source jobs.
2. **Compatibility Matrix** — filters on all six vendor CSVs and any exactly matched Broadcom scraped rows; source provenance, Excel download.
3. **PTF Intelligence** — product/release selection, extracted PTF IDs, linked PTF and Additional Information text, OS comparison, exact source links, Excel download.
4. **Source Monitor** — manual Broadcom refresh, option to follow approved off-page references, extraction metrics, errors and source directory.
5. **Change History** — compare successive successful Broadcom snapshots; inspect before/after records and export to Excel.
6. **Inventory Readiness** — upload customer CSV/XLSX with `Vendor, Product, VRM`, or select entries; view compatibility and PTF caveats and export results. Installed SMP/E service cannot be inferred.
7. **AI Assistant** — TF-IDF matching from all compatibility records; optional Granite text generation when IBM watsonx credentials are configured.

## Using the real Broadcom crawler

Open **Source Monitor → Refresh Broadcom now**. The backend:

1. Downloads the official Broadcom z/OS compatibility HTML matrix over HTTPS.
2. Finds the product table using its *z/OS V3R2, V3R1, V2R5* headers rather than assuming a fixed table index.
3. Extracts vendor, product, release, target OS, compatibility, and every PTF/Additional Information link associated with that OS cell.
4. Resolves same-page `#anchor` targets in the downloaded HTML. If the optional **Follow approved external detail pages** toggle is enabled, requests supported off-page HTTPS vendor links; only allowlisted vendor domains may be contacted.
5. Extracts source text, PTF/APAR IDs, FIXCAT references and possible USERMOD/SYSMOD/FMID indicators. Keeps the raw extracted text and original URL for manual audit.
6. Rejects obviously incomplete results, then atomically saves a new SQLite snapshot. The previous successful snapshot stays available on failure.
7. Replaces a baseline support status **only when vendor + product + release + z/OS match exactly after normalization**; other existing CSV rows are preserved. New source releases appear separately.

This is a conservative HTML parser, **not** a universal web crawler. Vendors may change page layouts, use authenticated portals, or include JS-driven or unresolvable links. Those links are marked partial/unavailable rather than fabricated. A live crawl must run in a network-connected environment; it is not run by packaging or tests.

**Important:** PTF identification is not PTF installation verification. IDs extracted from unrestricted narrative text may include related PTFs; read the original source to establish whether every ID is required. Source data is not a replacement for product lifecycle or SMP/E checks.

## Data and persistence

- `data/*_compat.csv`: original six baseline vendor datasets. No changes to source CSVs.
- `reference/zOS_3.2_filtered.xlsx`: the provided reference workbook.
- `state/compatibility.sqlite3`: created at first use and holds real crawled snapshots, per-OS maintenance evidence, change records and refresh job logs. It is **not** included in the distribution.
- `ZOS_DB_PATH` environment variable: overrides the SQLite path. E.g. `ZOS_DB_PATH=/data/compatibility.sqlite3` for a mounted persistent volume.

**Streamlit Community Cloud** and other ephemeral containers may erase local SQLite files during redeployment/restart. Use a persistent mounted volume or migrate the store layer to a managed database for long-lived production audit history. The app does **not** run recurring background polling: click Refresh or schedule an external job with persistence and appropriate vendor rate limits. The UI displays the last successful snapshot, not guaranteed instantaneous vendor updates.

## AI configuration (optional)

The original `rag_engine.py` is preserved. Set `WATSONX_APIKEY`, `WATSONX_PROJECT_ID` (or `WATSONX_SPACE_ID`) and optionally `WATSONX_URL`, `WATSONX_MODEL_ID` in private environment secrets. Install the optional `ibm-watsonx-ai` library to enable text generation. Without this configuration, the **AI Assistant page works as evidence retrieval**.

Never commit tokens or API keys to the ZIP or source repository.

## Run tests

```bash
python -m unittest discover -s tests -v
python -m compileall -q .
```

Offline fixture tests validate product/OS mapping, internal anchor PTF extraction, atomic snapshot versioning and failed-job preservation. They **do not validate vendor network connectivity or a full browser session**.

## Deployment

A maintained `Dockerfile` serves Streamlit on port 8080 for IBM Cloud Code Engine. `legacy_app.py` preserves the original application as a reference. Existing deployment and presentation notes are included in `reference/`.

Use HTTPS, keep external vendor traffic bounded, respect vendor access policies, and verify licensing/usage terms before scheduling repeated crawls. Separate source adapters are needed to extract *live* detail pages for IBM, BMC, Rocket, Precisely and ASG; the dashboard already displays their CSV baseline records.
