import sys
import os

def main():
    logs = []
    logs.append(f"sys.frozen: {getattr(sys, 'frozen', False)}")
    logs.append(f"sys._MEIPASS: {getattr(sys, '_MEIPASS', 'NOT_FOUND')}")
    logs.append(f"sys.executable: {sys.executable}")
    logs.append(f"__file__: {__file__ if '__file__' in globals() else 'NOT_FOUND'}")
    logs.append(f"os.getcwd(): {os.getcwd()}")
    
    with open(os.path.expanduser("~/Downloads/test_paths.log"), "w") as f:
        f.write("\n".join(logs))

if __name__ == '__main__':
    main()
