import sys
import streamlit.web.cli as stcli

if __name__ == "__main__":
    sys.argv = ["streamlit", "run", "mvp_audit_engine.py", "--global.developmentMode=false"]
    sys.exit(stcli.main())