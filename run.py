import sys
from pathlib import Path

# Add the project root to Python path
project_root = Path(__file__).parent
sys.path.append(str(project_root))

# Import and run the app
from src.gui.main_window import MainApp

if __name__ == '__main__':
    MainApp().run() 