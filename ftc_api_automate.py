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
]

class FTCAPI:
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
        self.response_file = "responses_ftc.txt"
        with open(self.response_file, "w") as f:
            f.write(f"Response Log - Created at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n")
        print(f"Response file '{self.response_file}' created.")

        # Initialize job IDs list directly
        self.job_ids = []
        print("Job IDs list initialized.")

    def extract_job_ids(self, response):
        """Extract job IDs from a response string."""
        return re.findall(r'Job ID="([^"]+)"', response)

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
        print(f"Starting automated testing for all FTC commands")
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

        # Initialize the port number
        port = ""
        #Change this port number to the one you want to test (ftc or cluster)
        self.click_position("port number")
        self.click_position("cluster port")

        # Check which port is selected
        if self.positions["port number"] == self.positions["cluster port"]:
            port = "cluster"
        elif self.positions["port number"] == self.positions["ftc port"]:
            port = "ftc"
        else:
            port = "redundancy"

        # Update the "Get List of Job IDs" command section
        if "Initial Get List of Job IDs" in self.positions:
            self.click_position("dropdown")
            self.click_position("Initial Get List of Job IDs")
            self.click_position("Parameter")
            pyautogui.hotkey("ctrl", "a")
            pyautogui.hotkey("delete")

            self.click_position("API name")
            pyautogui.hotkey("ctrl", "a")
            pyautogui.hotkey("delete")
            self.click_position("issue")
            response = self.copy_response()
            self.job_ids = self.extract_job_ids(response)
            print(f"Updated Job IDs: {self.job_ids}")
        else:
            print("Position for 'Initial Get List of Job IDs' not found in new_pos.json!")

        # Go back to the main dropdown
        self.click_position("dropdown")

        # Press down arrow several times to reach "Get System Info"
        for _ in range(9): 
            pyautogui.press('up')

        # Press Enter to select
        pyautogui.press('enter')


        # Variable to store the job ID of the "Add 2 min Fake Job"
        fake_2min_id = None

        # Variable to store the job ID of the "Add 20 secs Fake Job"
        fake_20secs_id = None

        # Iterate through all commands and write responses to the main file
        for i, command in enumerate(commands):
            self.click_position("dropdown")
            self.click_position(command)

            # Check if the command exists in the positions dictionary
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue

            if command.lower() == "get list of job ids":
                self.click_position("issue")
                # Copy the response and update the job_ids list
                response = self.copy_response()
                self.job_ids = self.extract_job_ids(response)
                print(f"Updated Job IDs: {self.job_ids}")
                continue  # Skip the generic logic for this command

            if command.lower() == "add 2 mins fake job":
                self.click_position("issue")
                # Extract the job ID from the response
                response = self.copy_response()  # Get the response from the clipboard
                match = re.search(r'JobID="([^"]+)"', response)
                if match:
                    fake_2min_id = match.group(1)
                    print(f"Extracted Job ID for 'Add 2 min Fake Job': {fake_2min_id}")
                    continue
                else:
                    print(f"Failed to extract Job ID for 'Add 2 mins Fake Job'. Response might not match the expected format.")
                    continue

            if command.lower() == "add 20 secs fake job":
                self.click_position("issue")
                # Extract the job ID from the response
                response = self.copy_response()
                match = re.search(r'JobID="([^"]+)"', response)
                if match:
                    fake_20secs_id = match.group(1)
                    print(f"Extracted Job ID for 'Add 20 secs Fake Job': {fake_20secs_id}")
                    continue
                else:
                    print(f"Failed to extract Job ID for 'Add 20 secs Fake Job'. Response might not match the expected format.")
                    continue

            # Check if the command requires a job ID
            if command.lower() == "remove job":
                # Ensure the job ID is different from fake_2min_id and fake_20secs_id
                job_id = next((jid for jid in self.job_ids if jid != fake_2min_id and jid != fake_20secs_id), None)
                if not job_id:
                    print(f"No suitable Job ID available for command '{command}'")
                    continue
                # Move cursor to the job ID field and paste the job ID
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(job_id)
                self.click_position("issue")
                self.write_response_to_file(self.response_file, command_name=command)
                self.job_ids.remove(job_id)  # Remove the used job ID from the list
                continue  # Skip the generic logic for this command

            # Special handling for "fake_20secs_id" job ID
            if command.lower() in ["get information about a job", "set job priority", "set job retries", "requeue a job", "cancel a job"]:
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(fake_20secs_id)
                self.click_position("issue")
                self.write_response_to_file(self.response_file, command_name=command)
                continue  # Skip the generic logic for this command

            # Special handling for "fake_2min_id" job ID
            if command.lower() in ["pause a job", "resume a job"]:

                # Move cursor to the job ID field and paste the job ID
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                pyautogui.typewrite(fake_2min_id)

                # Click the "Issue" button
                self.click_position("issue")

                # Write the response to the file
                self.write_response_to_file(self.response_file, command_name=command)
                continue

            # Generic logic for all other commands
            self.click_position("issue")
            self.write_response_to_file(self.response_file, command_name=command)

# Initialize and run the tests
ftc_automate = FTCAPI()
ftc_automate.run_tests()