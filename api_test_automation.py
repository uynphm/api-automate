import pyautogui
import time
import logging
import json
import pygetwindow as gw
import pyperclip  # Import pyperclip to handle clipboard operations
import re  # Import regex for extracting job IDs

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
    "Requeue a Job",
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
    "Set Number of Transcoding Slots to 2",
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
        pyautogui.PAUSE = 0.5
        # Enable fail-safe (move mouse to corner to stop)
        pyautogui.FAILSAFE = True

        # Load positions from JSON file
        try:
            with open('new_pos.json', 'r') as f:
                self.positions = json.load(f)
                print("Successfully loaded positions from new_pos.json")
        except Exception as e:
            logging.error(f"Error loading new_pos.json: {str(e)}")
            raise

        # Create a single response file at the start of the test
        self.response_file = "responses.txt"
        with open(self.response_file, "w") as f:
            f.write(f"Response Log - Created at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n")
        print(f"Response file '{self.response_file}' created.")

        # Extract job IDs from the job_ids.txt file
        self.job_ids = self.extract_job_ids("job_ids.txt")
        print(f"Extracted Job IDs: {self.job_ids}")

    def extract_job_ids(self, filename):
        """Extract job IDs from the job_ids.txt file."""
        job_ids = []
        try:
            with open(filename, "r") as f:
                content = f.read()
                # Use regex to find all job IDs in the file
                job_ids = re.findall(r'Job ID="([^"]+)"', content)
        except FileNotFoundError:
            logging.warning(f"File '{filename}' not found. No job IDs extracted.")
        return job_ids

    def click_position(self, label, delay=1):
        x, y = self.positions[label]
        pyautogui.moveTo(x, y)
        pyautogui.click()
        time.sleep(delay)

    def copy_response(self):
        self.click_position("response_box")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)  # Give it a moment
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)  # Give it a moment
        # Get the copied response from the clipboard
        response = pyperclip.paste()
        print(response)  # Print the response to the console
        return response

    def write_response_to_file(self, filename, command_name=""):
        # Copy the response and write to the specified file
        response = self.copy_response()  # Capture the response
        with open(filename, "a") as f:  # Open in append mode
            f.write(f"Command: {command_name}\n")
            f.write(response)  # Write the captured response
            f.write("\n" + "-" * 50 + "\n")  # Add a separator for readability
        print(f"Response for '{command_name}' written to {filename}")

    def run_tests(self):
        print(f"Starting automated testing for all commands")
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

        # Renew the "job_ids.txt" file
        job_ids_file = "job_ids.txt"
        with open(job_ids_file, "w") as f:
            f.write(f"Job IDs Log - Created at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n")
        print(f"File '{job_ids_file}' has been renewed.")

        # Issue the "Get List of Job IDs" command and save to the renewed file
        if "Initial Get List of Job IDs" in self.positions:
            self.click_position("dropdown")
            self.click_position("Initial Get List of Job IDs")
            self.click_position("issue")
            time.sleep(0.2)
            self.write_response_to_file(filename=job_ids_file, command_name="Initial Get List of Job IDs")
        else:
            print("Position for 'Initial Get List of Job IDs' not found in new_pos.json!")

        # Update the job_ids list with the latest data from the renewed file
        self.job_ids = self.extract_job_ids(job_ids_file)
        print(f"Updated Job IDs: {self.job_ids}")

        # Go back to the main dropdown
        self.click_position("dropdown")

        # Press down arrow several times to reach "Get System Info"
        for _ in range(9): 
            pyautogui.press('up')

        # Press Enter to select
        pyautogui.press('enter')


        # Variable to store the job ID of the "2 min Fake Job"
        fake_job_id = None

        # Iterate through all commands and write responses to the main file
        for i, command in enumerate(commands):
            self.click_position("dropdown")
            self.click_position(command)

            # Check if the command exists in the positions dictionary
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue

            if command == "Get List of Job IDs":
                self.click_position("issue")
                time.sleep(0.2)
                # Copy the response and update the job_ids_file
                response = self.copy_response()
                with open(job_ids_file, "w") as f:
                    f.write(f"Command: {command}\n")
                    f.write(response)
                    f.write("\n" + "-" * 50 + "\n")
                print(f"Response for '{command}' written to {job_ids_file} (file renewed).")

                # Update the job_ids list with the latest data from the updated file
                self.job_ids = self.extract_job_ids(job_ids_file)
                print(f"Updated Job IDs: {self.job_ids}")
                time.sleep(0.2)
                continue  # Skip the generic logic for this command

            if command == "Add 2 mins Fake Job":
                self.click_position("issue")
                # Extract the job ID from the response
                response = self.copy_response()  # Get the response from the clipboard
                match = re.search(r'JobID="([^"]+)"', response)
                print(match)
                if match:
                    fake_job_id = match.group(1)
                    print(f"Extracted Job ID for '2 min Fake Job': {fake_job_id}")
                    continue
                else:
                    print(f"Failed to extract Job ID for '2 min Fake Job'. Response might not match the expected format.")
                    continue

            # Check if the command requires a job ID
            if "get information about a job" in command.lower() or "set job priority" in command.lower() or "set job retries" in command.lower() or "remove job" in command.lower() or "requeue a job" in command.lower() or "cancel a job" in command.lower():
                # Ensure the job ID is different from fake_job_id
                job_id = next((jid for jid in self.job_ids if jid != fake_job_id), None)
                if not job_id:
                    print(f"No suitable Job ID available for command '{command}'")
                    continue

                print(f"Using Job ID '{job_id}' for command '{command}'")
                # Move cursor to the job ID field and paste the job ID
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(job_id)
                time.sleep(0.2)
                self.click_position("issue")
                time.sleep(0.2)
                self.write_response_to_file(self.response_file, command_name=command)
                time.sleep(2)
                continue  # Skip the generic logic for this command

            # Special handling for "Pause a Job" and related commands
            if command in ["Pause a Job", "Resume a Job"]:

                # Move cursor to the job ID field and paste the job ID
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(fake_job_id)
                time.sleep(0.2)

                # Click the "Issue" button
                self.click_position("issue")
                time.sleep(0.2)

                # Write the response to the file
                self.write_response_to_file(self.response_file, command_name=command)
                time.sleep(0.2)
                job_id = fake_job_id
                continue

            # Generic logic for all other commands
            self.click_position("issue")
            time.sleep(0.2)
            self.write_response_to_file(self.response_file, command_name=command)
            time.sleep(0.2)

# Initialize and run the tests
automation = APITestAutomation()
automation.run_tests()