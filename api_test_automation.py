import pyautogui
import time
import logging
import json
import pygetwindow as gw

# List of command names in the order they appear in the dropdown
commands = [
    "Get System Info",
    "Get System Info (Invalid URI)",
    "Get System Status",
    "Add 20 secs Fake Job",
    "Add 2 mins Fake Job",
    "Add Source Analysis Job",
    "Add Job with non-XML data",
    "Add Job with missing output location",
    "Remove Job",
    "Get List of Job IDs",
    "Get List of paused jobs' Job IDs",
    "Get Information about a Job",
    "Pause a Job",
    "Resume a Job",
    "Request a Job",
    "Cancel a Job",
    "Set Job Priority to 1",
    "Set Job Priority to 10",
    "Set Job Retries to 2",
    "Set Job Retries to 0",
    "Set Invalid Job Status",
    "Get Job IDs By Source Filename",
    "Get Job IDs By Output Filename",
    "Search Jobs",
    "Get Transcoding Slot Info",
    "Set Number of Transcoding Slots to 1",
    "Set Number of Transcoding Slots to 10",
    "Enable 'Pause other jobs for Urgent jobs'",
    "Disable 'Pause other jobs for Urgent jobs'",
    "Get Machine List (Cluster Only)",
    "Add Machine Details (Cluster Only)",
    "Connect Machine to Cluster (Cluster Only)",
    "Add New Machine to Cluster by IP (Cluster Only)",
    "Disconnect Machine from Cluster (Cluster Only)",
    "Disconnect (Maintenance mode) Machine from Cluster (Cluster Only)",
    "Remove Machine from Cluster by ID (Cluster Only)",
    "Reset Machine Counter (Cluster Only)",
    "Set Machine Transcoding Slots (Cluster Only)",
    "Get Machine Tags (Cluster Only)",
    "Set Machine Tags (Cluster Only)",
    "Get Machine Priority (Cluster Only)",
    "Set Machine Priority (Cluster Only)",
    "Get Replication Status (Cluster Only)",
    "Set Replication Status (Cluster Only)",
    "Custom PUT"
]

class APITestAutomation:
    def __init__(self):
        # Add a small delay between actions to make them more reliable
        pyautogui.PAUSE = 1
        # Enable fail-safe (move mouse to corner to stop)
        pyautogui.FAILSAFE = True

        # Load positions from JSON file
        try:
            with open('new_pos.json', 'r') as f:
                self.positions = json.load(f)
                logging.info("Successfully loaded positions from new_pos.json")
        except Exception as e:
            logging.error(f"Error loading new_pos.json: {str(e)}")
            raise

    def click_position(self, label, delay=1):
        x, y = self.positions[label]
        pyautogui.moveTo(x, y)
        pyautogui.click()
        time.sleep(delay)

    def scroll_down(self, times=7, delay=0.5):
        for _ in range(times):
            pyautogui.scroll(-1000)
            time.sleep(delay)

    def run_tests(self):
        logging.info(f"Starting automated testing for all commands")
        # Focus and resize the window
        windows = gw.getWindowsWithTitle("Capella FTC API Test tool")
        if not windows:
            print("Window not found!")
            exit()

        # Use the first match
        win = windows[0]

        # Bring the window to front
        if win.isMinimized:
            win.restore()
        win.activate()
        time.sleep(0.5)  # Give it a moment

        # Resize and move it to a fixed position
        win.moveTo(100, 100)     # Top-left corner (x, y)
        win.resizeTo(1000, 700)  # Width, height

        # First 23 commands (no scroll)
        for i, command in enumerate(commands):
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue
            if i == 0:
                self.click_position("issue")
            elif i < 23:
                self.click_position("dropdown")
                self.click_position(command)
                self.click_position("issue")
            elif i == 23:
                self.click_position("dropdown")
                self.scroll_down(7)
                self.click_position(command)
                self.click_position("issue")
            else:
                self.click_position("dropdown")
                self.click_position(command)
                self.click_position("issue")
            time.sleep(2)


# # Configure logging
# logging.basicConfig(
#     level=logging.INFO,
#     format='%(asctime)s - %(levelname)s - %(message)s',
#     handlers=[
#         logging.FileHandler(f'api_test_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
#         logging.StreamHandler()
#     ]
# )

# Initialize and run the tests
automation = APITestAutomation()
automation.run_tests()

# import pygetwindow as gw
# print([win.title for win in gw.getAllWindows()])