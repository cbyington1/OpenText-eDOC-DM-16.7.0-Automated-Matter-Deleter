import pyautogui
import win32gui
import win32con
import os
import time
import pygetwindow as gw
import psutil
import sys
import subprocess

def execute_application(application_path):
    os.startfile(application_path)

def find_window_by_date():
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if "1/1/1992" in window_title:
                hwnds.append(hwnd)
            elif any(char.isdigit() for char in window_title) and len(window_title) == 5:
                hwnds.append(hwnd)
        return True

    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows[0] if windows else None

def restart_application(matter_number, restartAttempts):
    # Kill any existing eDOCS DM processes
    for proc in psutil.process_iter(['pid', 'name']):
        if 'eDOCS DM' in proc.info['name']:
            print(f"Terminating eDOCS DM process with PID {proc.info['pid']}...")
            try:
                proc.terminate()
            except psutil.NoSuchProcess:
                pass

    # Close windows with the name "eDOCS DM"
    eDOCS_DM_windows = gw.getWindowsWithTitle("eDOCS DM")
    for window in eDOCS_DM_windows:
        window.close()

    os.system("taskkill /F /IM explorer.exe")
    if restartAttempts == 0:
        time.sleep(5)
        os.system("start explorer.exe")
        time.sleep(10)
    elif restartAttempts == 1:
        time.sleep(60)
        os.system("start explorer.exe")
        time.sleep(30)
    elif restartAttempts == 2:
        time.sleep(120)
        os.system("start explorer.exe")
        time.sleep(60)
    elif restartAttempts == 3:
        time.sleep(180)
        os.system("start explorer.exe")
        time.sleep(90)
    else:
        # For restartAttempts greater than 3, use the same values as for restartAttempts == 3
        time.sleep(180)
        os.system("start explorer.exe")
        time.sleep(90)

    extension_folder = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Open Text\\DM Extensions"
    for root, _, files in os.walk(extension_folder):
        for file in files:
            if file.endswith(".lnk"):
                execute_application(os.path.join(root, file))
                time.sleep(15)
                process_single_matter(matter_number, restartAttempts + 1)  # Restart the matter searching process
                # No need to return here, it will continue with the next iteration of the loop

def detect_internal_error(matter_number, restartAttempts):
    internal_error_window_title = "DOCSList_x64"  # Replace with the actual title of the error window
    internal_error_window = detect_popup_window_by_title(internal_error_window_title)

    if internal_error_window is not None:
        print("Internal error detected.")
        restart_application(matter_number, restartAttempts)

def search_process(matter_number, restartAttempts, attempts = 0):  # Move the search_process function outside of process_matters
    if(attempts > 2):
        print("Matter number field not found. Restarting application...\n")
        restart_application(matter_number, restartAttempts)
    detect_internal_error(matter_number, restartAttempts)
    search_button_location = pyautogui.locateOnScreen('search_button.png')
    if search_button_location is not None:
        search_button_center = pyautogui.center(search_button_location)
        pyautogui.click(search_button_center)

        time.sleep(5)

        matter_number_location = pyautogui.locateOnScreen('matter_number_field.png')
        if matter_number_location is not None:
            matter_number_center = pyautogui.center(matter_number_location)
            pyautogui.click(matter_number_center)
            pyautogui.typewrite(matter_number)
            pyautogui.move(250, 110)
            pyautogui.click()
            pyautogui.typewrite("1/1/1992 to 1/1/2017")

            pyautogui.press('enter')
            return

        if matter_number_location is None:
            time.sleep(10)
            matter_number_location = pyautogui.locateOnScreen('matter_number_field.png')
            if matter_number_location is not None:
                matter_number_center = pyautogui.center(matter_number_location)
                pyautogui.click(matter_number_center)
                pyautogui.typewrite(matter_number)
                pyautogui.move(250, 110)
                pyautogui.click()
                pyautogui.typewrite("1/1/1992 to 1/1/2017")

                pyautogui.press('enter')
                return
            else:
                time.sleep(5)
                search_process(matter_number, restartAttempts, attempts + 1)


def find_max_amount_found_image(matter_number, restartAttempts):
    window_title = "eDOCS DM"
    max_amount_found_window = gw.getWindowsWithTitle(window_title)

    if max_amount_found_window:
        pyautogui.click(1000, 600)
        time.sleep(2)

        # Check again if the pop-up window with the name "eDOCS DM" still exists
        max_amount_found_window = gw.getWindowsWithTitle(window_title)
        if max_amount_found_window:
            print("Pop-up window still exists. Restarting the application...")
            restart_application(matter_number, restartAttempts)

        return True
    else:
        return False

def detect_popup_window_by_title(window_title):
    try:
        popup_window = gw.getWindowsWithTitle(window_title)[0]
        return popup_window
    except IndexError:
        return None

def confirm_deletion_window_handler(matter_number, restartAttempts, attempts=0):
    if attempts > 5:
        print("Confirm deletion hasn't loaded after multiple attempts. Restarting the application...\n")
        restart_application(matter_number, restartAttempts)

    confirm_deletion_window_title = "eDOCS DM"  # Replace with the actual title of the pop-up window
    confirm_deletion_window = detect_popup_window_by_title(confirm_deletion_window_title)
    
    loading_deletion_image_path = 'loading_deletion.png'
    loading_deletion_location = pyautogui.locateOnScreen(loading_deletion_image_path)

    if confirm_deletion_window:
        if loading_deletion_location:
            time.sleep(10)
            confirm_deletion_window_handler(matter_number, restartAttempts, attempts + 1)
        else:
            print("Confirm deletion found. Continuing with the process...")
            time.sleep(2)
            return
    else:
        print("Confirm deletion has not been found. Restarting the application...\n")
        restart_application(matter_number, restartAttempts)

def process_single_matter(matter_number, restartAttempts):
    print("\nProcessing matter number:", matter_number)

    if restartAttempts > 2:
        print(f"Skipping matter number {matter_number} due to excessive restart attempts.")
        with open("C:\\Users\\cbyington\\Documents\\HummingBirdPics\\failed_matter_numbers.txt", "a") as file:
            file.write(matter_number + ";")
    else:
        search_process(matter_number, restartAttempts)

        time.sleep(35)
        detect_internal_error(matter_number, restartAttempts)

        no_items_location = pyautogui.locateOnScreen('no_items.png')
        if no_items_location is not None:
            print("No items found for matter number:", matter_number, "\n")
            time.sleep(3)
            return

        detect_internal_error(matter_number, restartAttempts)
        max_amount_found = False

        matter_window = find_window_by_date()
        if matter_window:
            win32gui.ShowWindow(matter_window, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(matter_window)

            for _ in range(10):
                if win32gui.GetForegroundWindow() == matter_window:
                    break
                time.sleep(0.5)
            else:
                print("Window focus not set after multiple attempts. Restarting the application...")
                restart_application(matter_number, restartAttempts)

            window_placement = win32gui.GetWindowPlacement(matter_window)
            if window_placement[1] != win32con.SW_SHOWMAXIMIZED:
                window_placement = list(window_placement)
                window_placement[1] = win32con.SW_SHOWMAXIMIZED
                win32gui.SetWindowPlacement(matter_window, tuple(window_placement))

            detect_internal_error(matter_number, restartAttempts)

            max_amount_found = find_max_amount_found_image(matter_number, restartAttempts)

            ok_sign_location = None
            if max_amount_found:
                ok_sign_location = pyautogui.locateOnScreen('ok_sign.png')
            if ok_sign_location:
                ok_sign_center = pyautogui.center(ok_sign_location)
                pyautogui.click(ok_sign_center)

            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.moveTo(250, 200)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'a')

            sleep_time = 12 if max_amount_found else 10
            time.sleep(sleep_time)

            pyautogui.hotkey('ctrl', 'a')

            sleep_time = 6 if max_amount_found else 3
            time.sleep(sleep_time)

            detect_internal_error(matter_number, restartAttempts)

            pyautogui.moveTo(90, 80)
            pyautogui.click()

            time.sleep(1)

            pyautogui.move(0, 325)

            time.sleep(1)

            pyautogui.move(250, 0)
            pyautogui.click()

            sleep_time = 20 if max_amount_found else 10
            time.sleep(sleep_time)

            # Check for deletion confirmation pop-up window
            confirm_deletion_window_handler(matter_number, restartAttempts)

            pyautogui.click(800, 570)

            time.sleep(2)

            detect_internal_error(matter_number, restartAttempts)

            while True:
                ok_sign_location = pyautogui.locateOnScreen('ok_sign.png')
                if ok_sign_location is not None:
                    ok_sign_center = pyautogui.center(ok_sign_location)
                    pyautogui.click(ok_sign_center)
                    pyautogui.move(100, 0)
                    time.sleep(0.25)
                else:
                    sleep_time = 50 if max_amount_found else 25
                    time.sleep(sleep_time)

                    ok_sign_location = pyautogui.locateOnScreen('ok_sign.png')
                    if ok_sign_location is None:
                        break
            print("Matter number", matter_number, "has been deleted.\n")
        else:
            print("Window not found. Restarting the application.\n")
            restart_application(matter_number, restartAttempts)  # make it so that the restart application feature doesn't move onto the next number

        sleep_time = 40 if max_amount_found else 20
        time.sleep(sleep_time)
    
if __name__ == "__main__":
    subprocess.run(["python", "C:\\Users\\cbyington\\Documents\\HummingBirdPics\\matterGrabber.py"])

    matter_numbers_file = "matter_numbers.txt"
    with open(matter_numbers_file, 'r') as file:
        matter_numbers = []
        for line in file:
            matter_numbers.extend(line.strip().split(';'))

    # Store the initial count of matter numbers
    total_to_process = len(matter_numbers)

    # Record the starting time
    start_time = time.time()

    while True:
        # Create a copy of the matter_numbers list to iterate over
        for matter_number in matter_numbers[:]:
            if not any(char.isdigit() for char in matter_number):
                print("\nScript has processed all matter files")
                sys.exit()

            process_single_matter(matter_number, 0)
            time.sleep(15)

            # Remove the processed matter number from the matter_numbers list
            matter_numbers.remove(matter_number)

            # Write the updated matter numbers back to the file
            with open(matter_numbers_file, 'w') as file:
                lines = []
                current_line = []
                for i, matter in enumerate(matter_numbers, start=1):
                    current_line.append(matter)
                    if i % 20 == 0:
                        lines.append(';'.join(current_line))
                        current_line = []
                if current_line:
                    lines.append(';'.join(current_line))
                file.write('\n'.join(lines))

            # Calculate the amount of matter numbers processed
            processed_count = total_to_process - len(matter_numbers)
            print(f"Processed {processed_count} out of {total_to_process} matter numbers.")

            # Calculate and print the elapsed time in hours, minutes, and seconds
            elapsed_time = time.time() - start_time
            hours, rem = divmod(elapsed_time, 3600)
            minutes, seconds = divmod(rem, 60)
            print(f"Elapsed time: {int(hours)} hours, {int(minutes)} minutes, {int(seconds)} seconds.")

            # Check if all matter numbers have been processed
            if processed_count >= total_to_process:
                print("\nScript has processed all matter files")
                sys.exit()
