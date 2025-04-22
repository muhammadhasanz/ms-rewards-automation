# Setting up Microsoft Rewards Bot as a Background Service

This guide explains how to set up the `ms_rewards_bot_headless.py` script to run automatically in the background on a schedule using system services.

This allows the bot to perform tasks daily without requiring you to manually run the script or keep a terminal window open.

**Important:** This guide assumes you have already:
1. Cloned the bot repository.
2. Installed the Python dependencies (`pip install -r requirements.txt`).
3. **CRITICALLY:** Run the standard script (`ms_rewards_bot.py`) **once** to perform the initial manual login and create the persistent browser profile. The headless script will use this saved profile.

---

## Running on Linux (using systemd User Services)

Systemd user services are the standard, robust way to run background tasks for a specific user on modern Linux. They ensure the script runs reliably on schedule and can restart if needed.

*(Note: You mentioned using `nohup`. While `nohup` allows a script to continue after you close the terminal, it's not a full service manager. It won't automatically start the script after a reboot, won't automatically restart it if it crashes, and managing its status is less straightforward than with systemd. Systemd is the recommended approach for unattended daily runs.)*

We will create two small files:
1.  A `.service` file: Tells systemd *how* to run your script once.
2.  A `.timer` file: Tells systemd *when* to run that service file.

1.  **Find Paths:** You need the full path to your Python interpreter and the bot script's main directory.
    *   Open your terminal and go to the bot's main directory.
    *   Find Python: Type `which python3` (or `which python`). Copy the output path (e.g., `/usr/bin/python3`).
    *   Find Bot Directory: Type `pwd`. Copy the output path (e.g., `/home/your_user/ms_rewards_bot`).

2.  **Create Service Directory:** Make the folder for user services if it's not there:
    ```bash
    mkdir -p ~/.config/systemd/user/
    ```

3.  **Create the `.service` file:**
    Create a file named `ms-rewards-bot.service` inside `~/.config/systemd/user/`.
    Edit the file and paste the following, replacing `/path/to/python` and `/path/to/your/bot/directory` with the paths you found in step 1.

    ```ini
    [Unit]
    Description=Microsoft Rewards Bot headless run
    # Requires graphical session startup for browser interaction, even headless sometimes depends on display libs
    After=graphical-session-pre.target
    Wants=graphical-session-pre.target

    [Service]
    Type=oneshot # Runs the script once and exits
    WorkingDirectory=/path/to/your/bot/directory # Important for logs and profile
    ExecStart=/path/to/python /path/to/your/bot/directory/ms_rewards_bot_headless.py
    # Optional: Add --nosearch if you don't want searches
    # ExecStart=/path/to/python /path/to/your/bot/directory/ms_rewards_bot_headless.py --nosearch

    # Direct logging to journald for easy viewing with journalctl
    StandardOutput=journal
    StandardError=journal

    [Install]
    WantedBy=graphical-session.target
    ```

4.  **Create the `.timer` file:**
    Create a file named `ms-rewards-bot.timer` inside `~/.config/systemd/user/`.
    Edit the file and paste the following. Change `10:00` to your desired daily run time (24-hour format).

    ```ini
    [Unit]
    Description=Run Microsoft Rewards Bot daily timer

    [Timer]
    # Run the service daily at the specified time (e.g., 10:00 AM)
    OnCalendar=daily 10:00
    # Optional: Instead of a precise time, run sometime shortly after midnight
    # This spreads out runs and avoids hitting servers right at the start of the day.
    # Uncomment and adjust the line below if you prefer this:
    # OnCalendar=*-*-* 00:00..01:00
    # AccuracySec=1h # Combine with the above line to allow up to 1 hour jitter

    Persistent=true # Ensures the timer runs even if the computer was off

    [Install]
    WantedBy=timers.target
    ```
    *   Choose only **one** `OnCalendar` line.

5.  **Activate Systemd Files:**
    Tell systemd about the new files:
    ```bash
    systemctl --user daemon-reload
    ```
    Enable the timer so it starts every time you log in:
    ```bash
    systemctl --user enable ms-rewards-bot.timer
    ```
    Start the timer now (it will wait for the next scheduled time):
    ```bash
    systemctl --user start ms-rewards-bot.timer
    ```

6.  **Check It's Running:**
    See if the timer is active:
    ```bash
    systemctl --user status ms-rewards-bot.timer
    ```
    View logs (most recent first):
    ```bash
    journalctl --user -u ms-rewards-bot.service
    ```
    Also check the `ms_rewards_automation.log` file in your bot's directory.

**Important for Linux:** User systemd services are tied to your user session. If you log out completely, the timer and service might stop. To make them run even when you are not logged in graphically, you might need to enable "lingering" for your user: `loginctl enable-linger your_username` (replace `your_username`).

---

## Running on Windows (using Task Scheduler)

Windows Task Scheduler is the built-in tool to run programs automatically on a schedule.

1.  **Find Paths:** You need the full path to your Python program and the bot script's directory.
    *   Open **Command Prompt** or **PowerShell**.
    *   Find Python: Type `where python`. Copy the path (e.g., `C:\Users\your_user\AppData\Local\Programs\Python\Python39\python.exe`).
    *   Find Bot Directory: Navigate to your bot's main folder (`cd C:\path\to\your\bot`) and then type `cd`. Copy the output path (e.g., `C:\Users\your_user\Documents\GitHub\ms_rewards_bot`).

2.  **Open Task Scheduler:** Search for "Task Scheduler" in the Windows search bar and open it.

3.  **Create a Task:** In the right-hand panel, click "Create Task...". (Don't use "Create Basic Task...")

4.  **General Tab:**
    *   **Name:** Type `Microsoft Rewards Bot`.
    *   **Description:** Add `Automates Microsoft Rewards tasks daily`.
    *   **Security options:**
        *   Select **"Run whether user is logged on or not"**. This is key for it to run in the background.
        *   You will be asked for your Windows username and password. You **must** provide them for this option to work.
        *   Make sure "Do not store password" is **unchecked**.
        *   Check "Run with highest privileges" (Helps avoid permission problems).
        *   Configure for: Select your Windows version.

5.  **Triggers Tab:**
    *   Click "New...".
    *   **Begin the task:** Select "On a schedule".
    *   **Settings:** Choose "Daily".
    *   **Start:** Set the date and the daily **time** you want it to run (e.g., 10:00:00 AM).
    *   Make sure "Enabled" is checked. Click "OK".

6.  **Actions Tab:**
    *   Click "New...".
    *   **Action:** Select "Start a program".
    *   **Settings:**
        *   **Program/script:** Enter the full path to your Python program (e.g., `C:\Users\your_user\AppData\Local\Programs\Python\Python39\python.exe`). Use the path from Step 1.
        *   **Add arguments (optional):** Enter the full path to your `ms_rewards_bot_headless.py` script.
            *   Example: `C:\Users\your_user\Documents\GitHub\ms_rewards_bot\ms_rewards_bot_headless.py`
            *   To skip searches, add `--nosearch` after the script path: `C:\Users\your_user\Documents\GitHub\ms_rewards_bot\ms_rewards_bot_headless.py --nosearch`
        *   **Start in (optional):** Enter the full path to your bot's main directory (e.g., `C:\Users\your_user\Documents\GitHub\ms_rewards_bot`). This sets the working directory so the `ms_rewards_automation.log` file is created here. Use the path from Step 1.
    *   Click "OK".

7.  **Conditions & Settings Tabs:** You usually don't need to change anything here for a basic setup, but you can review them if you have specific needs (like only running on AC power).

8.  **Save Task:** Click "OK" to finish. You'll likely be prompted for your Windows password again to confirm the task creation.

9.  **Test It:**
    *   Find your new task (`Microsoft Rewards Bot`) in the Task Scheduler Library list.
    *   Right-click the task and select "Run".
    *   Check the "History" tab for the task (you might need to enable "Show History" in the Task Scheduler Actions panel first) to see if it completed.
    *   Check the `ms_rewards_automation.log` file in the directory you put in the "Start in" field of the task's action tab. This file will show the bot's progress and any errors.

**Important for Windows:** Running "whether user is logged on or not" requires storing your password with the task. Ensure your system is secure. Also, make sure your antivirus or other security software doesn't block the script, Python, Edge, or the WebDriver.