import os
import streamlit.web.cli as stcli
import sys

if __name__ == "__main__":
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Home_Page.py")

    if not os.path.exists(app_path):
        print(f"Error: {app_path} not found")
        sys.exit(1)

    print("Press Ctrl+C to stop Streamlit...")
    
    sys.argv = [
        "streamlit", 
        "run", app_path,
        "--server.port=8501",
        "--global.developmentMode=false"  # Disable dev mode
    ]
    
    stcli.main()