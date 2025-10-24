from src.assistant import Assistant
from src.gui import AssistantGUI

def main():
    """
    Main function for the desktop assistant.
    """
    assistant = Assistant()
    gui = AssistantGUI(assistant)
    gui.start()

if __name__ == "__main__":
    main()
