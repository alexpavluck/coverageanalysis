Coverage Survey Analysis Tool (CSAT)

A browser-based tool for analysing ESPEN Coverage Evaluation Survey (CES) data and generating a standard post-MDA report. No server, no installation — the file never leaves your browser.

Usage


Open the tool in your browser
Fill in survey metadata (country, disease, drug, MDA round, thresholds)
Upload your ESPEN CES response data file (.xlsx, sheet name data)
Enter administratively-reported coverage per district/IU if available (optional — used to validate against survey results)
Click Generate report


The report renders inline and can be saved as a PDF via the Save as PDF button.

Data format

Expects an ESPEN Coverage Evaluation Survey response file with the standard p_* column schema (e.g. p_district, p_site, p_received_ivm, p_swalllowed_ivm). Supports IVM, ALB, and PZQ drug codes.

What the report includes


Epidemiological and therapeutic coverage by district, with 95% cluster-robust logit CIs
Coverage by site (cluster-level dot plots)
Treatment pathway cascade (Total → Received → Swallowed)
Never-treated population analysis
Side effects, communication reach, and community satisfaction
Comparison of survey results against reported coverage (if provided)
WHO interpretation framework


Technical notes


Runs entirely in the browser using Pyodide (Python via WebAssembly)
No data is uploaded to any server
First load downloads ~10 MB of the Pyodide runtime; subsequent loads use the browser cache
Requires an internet connection to load dependencies (Pyodide, SheetJS) from CDN
