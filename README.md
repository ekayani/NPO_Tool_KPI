# RAN KPI Analytics Tool (v13)

## Run locally
    pip install -r requirements.txt
    streamlit run ran_kpi_analytics_tool_v13.py

## Deploy on Streamlit Community Cloud
1. Create a GitHub repo and push these files (keep your KPI/planning data OUT of the repo):
       ran_kpi_analytics_tool_v13.py, requirements.txt, .streamlit/config.toml, .gitignore, README.md
2. Go to share.streamlit.io, sign in with GitHub, click "Create app".
3. Pick the repo, branch `main`, main file `ran_kpi_analytics_tool_v13.py`.
4. Advanced settings: Python 3.11 or 3.12. Paste your secrets (see .streamlit/secrets.toml.example).
5. Deploy. Later pushes to the branch redeploy automatically.

## Access control
Public by default. In the app's Settings > Sharing you can restrict viewing to
specific email addresses, recommended since users upload operator network data.

## Limits to know
Free tier has about 1 GB RAM. Very large KPI exports may need to be filtered
or split before upload. Uploads are held in memory per session, not stored.
"# NPO_Tool_KPI" 
