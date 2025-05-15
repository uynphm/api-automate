import pyautogui
import time
import logging
import json
import pygetwindow as gw
import pyperclip  # Import pyperclip to handle clipboard operations
import re  # Import regex for extracting job IDs

commands = [
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

redundancy_commands = [
    "Set Replication Status (Cluster Only)",
    "Get Replication Status (Cluster Only)"
]

class ClusterAPI:
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
        self.response_file = "responses_cluster.txt"
        with open(self.response_file, "w") as f:
            f.write(f"Response Log - Created at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n")
        print(f"Response file '{self.response_file}' created.")

        # Initialize machine IDs list directly
        self.machine_ids = []
        print("Machine IDs list initialized.")

    def extract_machine_ids(self, response):
        """Extract machine IDs from a response string."""
        return re.findall(r'Identifier="([^"]+)"', response)
    
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
        print(f"Starting automated testing for all Cluster commands")
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

        # Click the "Port Number" tab to change the port number to Cluster
        self.click_position("port number")
        time.sleep(0.1)  # Give it a moment
        self.click_position("cluster port")
        time.sleep(0.1)  # Give it a moment

        # Issue the "Get Machine List (Cluster Only)" command
        self.click_position("dropdown")
        for _ in range(29):
            pyautogui.press("down")
        pyautogui.press("enter")
        self.click_position("issue")

        self.write_response_to_file(self.response_file, "Get Machine List (Cluster Only)")
        response = self.copy_response()
        self.machine_ids = self.extract_machine_ids(response)
        print(f"Extracted Machine IDs: {self.machine_ids}")
        
        # Start enumeration for commands
        for i, command in enumerate(commands):
            try:
                self.click_position("dropdown")
                self.click_position(command)
            except KeyError:
                logging.error(f"Position for command '{command}' not found in new_pos.json!")
                continue

            # Check if the command exists in the positions dictionary
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue

            elif command == "Add New Machine to Cluster by IP (Cluster Only)":
                # Click the "Add New Machine to Cluster by IP" button
                self.click_position("Add New Machine to Cluster by IP (Cluster Only)")

                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")

                # Write a simulated IP address to the parameter field
                pyautogui.write("192.168.1.100")

                self.click_position("issue")
                self.write_response_to_file(self.response_file, command)

                # Extract the machine ID from the response
                response = self.copy_response()
                new_machine_id = re.search(r'MachineID="([^"]+)"', response)
                if new_machine_id:
                    self.machine_ids.append(new_machine_id.group(1))
                else:
                    logging.error("Failed to extract machine ID from Add New Machine response.")
                continue

            elif command == "Remove Machine from Cluster by ID (Cluster Only)":
                # Click the "Remove Machine from Cluster by ID" button
                self.click_position("Remove Machine from Cluster by ID (Cluster Only)")
                self.click_position("Job ID")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")

                # Use the machine ID extracted from the Add New Machine command
                if self.machine_ids:
                    machine_id = self.machine_ids.pop(-1)  # Remove the last machine ID added
                    pyautogui.write(machine_id)
                else:
                    logging.error("No machine IDs available to remove.")
                    continue

                self.click_position("issue")
                self.write_response_to_file(self.response_file, command)
                continue

            elif command in ["Set Machine Transcoding Slots (Cluster Only)", "Set Machine Tags (Cluster Only)", "Set Machine Priority (Cluster Only)"]:
                self.click_position(command)
                self.click_position("API name")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")
                if command == "Set Machine Transcoding Slots (Cluster Only)":
                    # Write a simulated number of slots to the parameter field
                    pyautogui.write("30")
                elif command == "Set Machine Tags (Cluster Only)":
                    # Write a simulated tag to the parameter field
                    pyautogui.write("Test,Tag")
                else:
                    # Write a simulated priority to the parameter field
                    pyautogui.write("3")
                self.click_position("issue")
                self.write_response_to_file(self.response_file, command)
                continue
            else:
                self.click_position("Parameter")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")

                # Use the first machine ID for all other commands
                if self.machine_ids:
                    pyautogui.write(self.machine_ids[0])  # Use the first machine ID without removing it
                else:
                    logging.error("No machine IDs available for the command.")
                    continue

                self.click_position("issue")
                self.write_response_to_file(self.response_file, command)

        # Handle redundancy commands
        self.click_position("port number")
        self.click_position("redundancy port")
        for i, command in enumerate(redundancy_commands):
            try:
                self.click_position("dropdown")
                self.click_position(command)
            except KeyError:
                logging.error(f"Position for command '{command}' not found in new_pos.json!")
                continue

            # Check if the command exists in the positions dictionary
            if command not in self.positions:
                logging.warning(f"Position for command '{command}' not found in new_pos.json!")
                continue

            elif command == "Set Replication Status (Cluster Only)":
                # Click the "Set Replication Status" button
                self.click_position("Set Replication Status (Cluster Only)")
                self.click_position("Parameter")
                pyautogui.hotkey("ctrl", "a")
                pyautogui.hotkey("delete")

                # Write a simulated replication status to the parameter field
                pyautogui.write("?Role=nobackup")
                self.click_position("issue")
                self.write_response_to_file(self.response_file, command)
                continue
            else:
               self.click_position("issue")
               self.write_response_to_file(self.response_file, command)

# Initialize and run the ClusterAPI
cluster_automate = ClusterAPI()
cluster_automate.run_tests()