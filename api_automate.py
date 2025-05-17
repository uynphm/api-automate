import pyautogui
import time
import logging
import json
import pygetwindow as gw
import pyperclip
import re
import sys

class APIAutomate:
    def __init__(self, mode):
        self.mode = mode.lower()
        pyautogui.PAUSE = 0.5
        pyautogui.FAILSAFE = True
        try:
            with open('new_pos.json', 'r') as f:
                self.positions = json.load(f)
                print("Successfully loaded positions from new_pos.json")
        except Exception as e:
            logging.error(f"Error loading new_pos.json: {str(e)}")
            raise

        # Define FTC and Cluster command lists for use in all modes
        self.ftc_commands = [
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
            "Requeue a Job",
            "Cancel a Job",
            "Set Job Priority to 1",
            "Set Job Priority to 10",
            "Set Job Retries to 0",
            "Set Job Retries to 2",
            "Set Invalid Job Status",
            "Get Job IDs By Source Filename",
            "Get Job IDs By Output Filename",
            "Search Jobs",
            "Get Transcoding Slot Info",
            "Set Number of Transcoding Slots to 1",
            "Set Number of Transcoding Slots to 2",
            "Enable 'Pause other jobs for Urgent jobs'",
            "Disable 'Pause other jobs for Urgent jobs'",
        ]
        self.cluster_commands = [
            "Get Machine List (Cluster Only)",
            "Get Machine Details (Cluster Only)",
            "Connect Machine to Cluster (Cluster Only)",
            "Add New Machine to Cluster by IP (Cluster Only)",
            "Disconnect Machine from Cluster (Cluster Only)",
            "Disconnect (Maintenance mode) Machine from Cluster (Cluster Only)",
            "Remove Machine from Cluster by ID (Cluster Only)",
            "Reset Machine Counter (Cluster Only)",
            "Set Machine Transcoding Slots (Cluster Only)",
            "Set Machine Tags (Cluster Only)",
            "Get Machine Tags (Cluster Only)",
            "Set Machine Priority (Cluster Only)",
        ]
        self.redundancy_commands = [
            "Set Replication Status (Cluster Only)",
            "Get Replication Status (Cluster Only)"
        ]

        if self.mode == "ftc":
            self.response_file = "responses.txt"
            self.job_ids = []
        elif self.mode == "cluster":
            self.response_file = "responses.txt"
            self.machine_ids = []
        else:
            raise ValueError("Mode must be 'ftc' or 'cluster'")
        with open(self.response_file, "w") as f:
            f.write(f"Response Log - Created at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n")
        print(f"Response file '{self.response_file}' created.")

    def click_position(self, label, delay=1):
        x, y = self.positions[label]
        pyautogui.moveTo(x, y)
        pyautogui.click()
        time.sleep(delay)

    def copy_response(self):
        self.click_position("response_box")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)
        response = pyperclip.paste()
        print(response)
        return response

    def write_response_to_file(self, command_name=""):
        response = self.copy_response()
        with open(self.response_file, "a") as f:
            f.write(f"Command: {command_name}\n")
            f.write(response)
            f.write("\n" + "-" * 50 + "\n")
        print(f"Response for '{command_name}' written to {self.response_file}")
        return response

    def extract_ids(self, response):
        # Support both JobID and Job ID
        matches = re.findall(r'JobID="([^"]+)"|Job ID="([^"]+)"', response)
        # Flatten and filter empty
        return [m[0] or m[1] for m in matches]

    def run_tests(self):
        print(f"Starting automated testing for {self.mode.upper()} commands")
        windows = gw.getWindowsWithTitle("Capella FTC API Test tool")
        if not windows:
            print("Window not found!")
            exit()
        win = windows[0]
        # Bring the window to front using alternative method if needed
        if win.isMinimized:
            win.restore()
        try:
            win.activate()
        except Exception as e:
            print(f"Warning: Could not activate window. {e}")
            # Alternative: try minimize and restore to force focus
            try:
                win.minimize()
                time.sleep(0.2)
                win.restore()
                time.sleep(0.2)
                win.activate()
            except Exception as e2:
                print(f"Alternative window focus also failed: {e2}")
        time.sleep(0.5)
        win.moveTo(100, 100)
        win.resizeTo(1000, 700)

        if self.mode == "ftc":
            # Do NOT change port, just run all FTC commands
            self.run_ftc_commands()
        elif self.mode == "cluster":
            # Change port to cluster, run all FTC commands, then all cluster and redundancy commands
            self.click_position("port number")
            self.click_position("cluster port")
            time.sleep(0.1)
            self.run_ftc_commands()
            self.run_cluster_commands()
            self.run_redundancy_commands()

    def run_ftc_commands(self):
        # Initialize job_ids
        self.job_ids = []
        fake_2min_id = None
        fake_20secs_id = None
        source_analysis_id = None
        for command in self.ftc_commands:
            self.click_position("dropdown")
            self.click_position(command)
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue
            if command == "Get List of Job IDs":
                self.click_position("issue")
                response = self.copy_response()
                self.job_ids = self.extract_ids(response)
                print(f"Updated Job IDs: {self.job_ids}")
                continue
            if command == "Add 2 mins Fake Job":
                self.click_position("issue")
                response = self.copy_response()
                match = re.search(r'JobID="([^"]+)"', response)
                if match:
                    fake_2min_id = match.group(1)
                    print(f"Extracted Job ID for 'Add 2 min Fake Job': {fake_2min_id}")
                continue
            if command == "Add 20 secs Fake Job":
                self.click_position("issue")
                response = self.copy_response()
                match = re.search(r'JobID="([^"]+)"', response)
                if match:
                    fake_20secs_id = match.group(1)
                    print(f"Extracted Job ID for 'Add 20 secs Fake Job': {fake_20secs_id}")
                continue
            if command == "Add Source Analysis Job":
                self.click_position("issue")
                response = self.copy_response()
                match = re.search(r'JobID="([^"]+)"', response)
                if match:
                    source_analysis_id = match.group(1)
                    print(f"Extracted Job ID for 'Add Source Analysis Job': {source_analysis_id}")
                continue
            if command == "Remove Job":
                # Prefer to use source_analysis_id if available
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(source_analysis_id)
                self.click_position("issue")
                self.write_response_to_file(command)
                if source_analysis_id in self.job_ids:
                    self.job_ids.remove(source_analysis_id)
                continue
            if command in ["Set Job Priority to 1", "Set Job Priority to 10", "Set Job Retries to 2", "Set Job Retries to 0", "Requeue a Job", "Cancel a Job", "Get Information about a Job"]:
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(fake_20secs_id)
                self.click_position("issue")
                self.write_response_to_file(command)
                continue
            if command in ["Pause a Job", "Resume a Job"]:
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(fake_2min_id)
                self.click_position("issue")
                self.write_response_to_file(command)
                continue
            self.click_position("issue")
            self.write_response_to_file(command)

    def run_cluster_commands(self):
        for command in self.cluster_commands:
            self.click_position("dropdown")
            self.click_position(command)
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue
            if command == "Get Machine List (Cluster Only)":
                self.click_position("issue")
                response = self.copy_response()
                machine_ids = re.findall(r'Identifier="([^"]+)"', response)
                if machine_ids:
                    self.machine_ids = machine_ids
                    print(f"Extracted Machine IDs: {self.machine_ids}")
                else:
                    logging.error("Failed to extract machine IDs from Get Machine List response.")
                continue
            if command == "Get Machine Details (Cluster Only)":
                self.click_position("Parameter")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                if self.machine_ids:
                    pyautogui.write(self.machine_ids[0])
                else:
                    logging.error("No machine IDs available for the command.")
                    continue
                self.click_position("issue")
                self.write_response_to_file(command)
                continue
            if command == "Add New Machine to Cluster by IP (Cluster Only)":
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.write("192.168.1.100")
                self.click_position("issue")
                self.write_response_to_file(command)
                response = self.copy_response()
                new_machine_id = re.search(r'MachineID="([^"]+)"', response)
                if new_machine_id:
                    self.machine_ids.append(new_machine_id.group(1))
                else:
                    logging.error("Failed to extract machine ID from Add New Machine response.")
                continue
            if command == "Remove Machine from Cluster by ID (Cluster Only)":
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                if self.machine_ids:
                    machine_id = self.machine_ids.pop(-1)
                    pyautogui.write(machine_id)
                else:
                    logging.error("No machine IDs available to remove.")
                    continue
                self.click_position("issue")
                self.write_response_to_file(command)
                continue
            if command in ["Set Machine Transcoding Slots (Cluster Only)", "Set Machine Tags (Cluster Only)", "Set Machine Priority (Cluster Only)"]:
                self.click_position("API name")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                if command == "Set Machine Transcoding Slots (Cluster Only)":
                    pyautogui.write("30")
                elif command == "Set Machine Tags (Cluster Only)":
                    pyautogui.write("Test,Tag")
                else:
                    pyautogui.write("3")
                self.click_position("issue")
                self.write_response_to_file(command)
                continue
            self.click_position("Parameter")
            pyautogui.hotkey("ctrl", "a")
            pyautogui.hotkey("delete")
            if self.machine_ids:
                pyautogui.write(self.machine_ids[0])
            else:
                logging.error("No machine IDs available for the command.")
                continue
            self.click_position("issue")
            self.write_response_to_file(command)

    def run_redundancy_commands(self):
        self.click_position("port number")
        self.click_position("redundancy port")
        for command in self.redundancy_commands:
            self.click_position("dropdown")
            self.click_position(command)
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue
            if command == "Set Replication Status (Cluster Only)":
                self.click_position("Parameter")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.write("?Role=nobackup")
                self.click_position("issue")
                self.write_response_to_file(command)
                continue
            self.click_position("issue")
            self.write_response_to_file(command)

if __name__ == "__main__":
    while True:
        mode = input("Which port do you want to test? Type 'ftc' or 'cluster': ").strip().lower()
        if mode in ("ftc", "cluster"):
            break
        print("Invalid input. Please type 'ftc' or 'cluster'.")
    automate = APIAutomate(mode)
    automate.run_tests()
