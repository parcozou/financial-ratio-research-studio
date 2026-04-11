# Financial Ratio Research Studio

## 1. Problem & User
This app is designed to support business students who need a simpler way to retrieve WRDS financial data, calculate key financial ratios, visualise trends, and generate short AI-supported interpretations. It turns a notebook-based workflow into a more accessible Streamlit interface for coursework, classroom demonstrations, and independent financial analysis practice.

## 2. Data
The current version supports only one WRDS source: **Compustat - North America**. The app retrieves company-level annual accounting and market fields such as company name, ticker, fiscal year, SIC code, sales, cost of goods sold, net income, EBIT, interest expense, depreciation, total assets, total liabilities, current assets, current liabilities, cash, short-term investments, receivables, inventory, long-term debt, equity, share price, and shares outstanding.

## 3. Methods
The workflow is organised into a small number of Python steps:

1. Collect user inputs in Streamlit, including ticker codes, fiscal-year range, optional SIC benchmark, WRDS credentials, and selected ratio metrics.
2. Connect to WRDS and retrieve annual company records from Compustat - North America.
3. Clean the data and calculate selected financial ratios, with fallback rules for missing fields where the calculation still remains financially reasonable.
4. Build formatted tables, compact markdown tables, and SVG-based charts for the requested metrics.
5. Optionally send the markdown tables to either LM Studio or an online OpenAI-compatible API for a short AI summary.
6. Return the results in a local or cloud-ready interface, with download options for tables, charts, and markdown outputs.

## 4. Key Findings
- The app supports two runtime modes in one codebase: `local` mode and `cloud` mode.
- `Local` mode supports WRDS login, optional `pgpass.conf`, LM Studio, online OpenAI-compatible APIs, and local export folders.
- `Cloud` mode is designed for Streamlit Community Cloud, uses browser downloads instead of server-side save paths, and uses online OpenAI-compatible APIs instead of localhost-only LM Studio.
- The app currently focuses only on **retrieving data and calculating financial ratios** rather than broader financial statement analysis, event studies, or multi-database research workflows.
- The AI summary prompt adjusts to the report structure, so it uses different language for a single-company case, a multi-company comparison, and a company-plus-industry-benchmark case.

## 5. How to run
### Local run
Install dependencies:

```bash
pip install -r requirements.txt
```

Start the app from the terminal:

```bash
streamlit run streamlit_app.py
```

Or double-click:

`Launch Financial Ratio Research Studio.bat`

### Streamlit Community Cloud deployment
Upload or push the main project files to GitHub, deploy `streamlit_app.py` as the entry file, and add the required secrets in the Streamlit Community Cloud secrets panel. In cloud mode, users should rely on the online OpenAI-compatible API option rather than LM Studio.

## 6. Product link / Demo
Streamlit Community Cloud product link:

[Financial Ratio Research Studio](https://financial-ratio-research-studio.streamlit.app/)

## 7. Limitations & Next steps
This version supports only **Compustat - North America** within WRDS, and it is limited to **retrieving company data and calculating financial ratios**. It does not yet support other WRDS databases, broader market datasets, non-ratio analytics, or more advanced research modules. If users request too many companies, years, and metrics at the same time, the AI summary step may take much longer to respond and may return an error without generating a result. The next development stage will expand to more databases, more analytical workflows, richer comparison logic, and broader export/reporting support.

## Appendix
### Local LM Studio + Gemma 4 workflow
If a user wants to run AI summaries locally instead of using an online API, the recommended workflow is:

1. Download and install LM Studio from [LM Studio Download](https://lmstudio.ai/download).
2. Open LM Studio and download a local Gemma model, for example [gemma-4-E4B-it-GGUF](https://huggingface.co/lmstudio-community/gemma-4-E4B-it-GGUF).
3. Load the model inside LM Studio.
4. Start LM Studio's local OpenAI-compatible server.
5. In the app, choose `LM Studio (local)` as the summary provider.
6. Keep the LM Studio server URL as `http://localhost:1234/v1` unless it has been changed manually.
7. Run the report and let the app send the compact markdown tables to the local Gemma model for a short summary.

### Project files
- `streamlit_app.py`: Streamlit user interface and runtime-mode logic.
- `financial_ratio_core.py`: WRDS retrieval, ratio calculation, export, and AI-summary helper functions.
- `requirements.txt`: Python dependency list.
- `Launch Financial Ratio Research Studio.bat`: one-click local launcher for Windows users.
