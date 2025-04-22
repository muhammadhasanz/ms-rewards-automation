# Microsoft Rewards Automation Bot

![Python](https://img.shields.io/badge/Python-3.6+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-%E2%99%A5-lightgrey.svg)
![Microsoft Edge](https://img.shields.io/badge/Browser-Edge-0078D4.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

This project provides a Python script to automate daily tasks on the Microsoft Rewards website using the Microsoft Edge browser via Selenium. It helps you earn points automatically by performing searches and completing available activities.

**Disclaimer:** This script is intended for personal use to automate repetitive tasks on your own Microsoft Rewards account. Use it responsibly and be aware of Microsoft's Terms of Service regarding automated activities. Using this script is at your own risk.

## Features

*   Automates Bing searches for both desktop and mobile points.
*   Identifies and attempts to complete Daily Set tasks (Quizzes, Polls, This or That, etc.).
*   Identifies and attempts to complete other available point-earning activities (Quizzes, surveys, video watches, etc.).
*   Uses a persistent Edge browser profile to maintain your login session and settings across runs.
*   Logs all activities and reports the points balance before and after the workflow.
*   Includes two versions: a standard visible browser version (`ms_rewards_bot.py`) and a headless background version (`ms_rewards_bot_headless.py`).
*   The headless version is designed to run once and exit, making it suitable for scheduling with system tools like `systemd` (Linux) or Task Scheduler (Windows).

## Prerequisites

*   Python 3.6 or higher.
*   Microsoft Edge Browser installed on your system.
*   An active Microsoft Account with access to the Microsoft Rewards program.
*   An internet connection.

## Installation

1.  **Clone the repository:**
    Open your terminal or command prompt and clone the project:
    ```bash
    git clone https://github.com/hei1sme/ms_rewards_bot.git
    cd ms_rewards_bot
    ```
    *(Replace `hei1sme` with your GitHub username and the repository name if you've forked it)*

2.  **Install Python dependencies:**
    Make sure you are in the `ms_rewards_bot` directory. The required libraries are listed in `requirements.txt`. Install them using pip:
    ```bash
    pip install -r requirements.txt
    ```
    This will install `selenium`, `webdriver-manager`, and `schedule`.

## Usage

There are two main scripts provided:

1.  `ms_rewards_bot.py`: This is the standard script that opens a visible Edge browser window when it runs. It's primarily used for the **initial login** and for interactive runs if you want to see the automation happening. It also includes an internal scheduler (though external scheduling is recommended for unattended use).
2.  `ms_rewards_bot_headless.py`: This is a modified version of the script designed to run in the background without opening a visible browser window. It executes the rewards tasks once when run and then exits. This version is ideal for scheduling with your operating system's task management tools.

---

### **🚨 FIRST RUN & MANUAL LOGIN (REQUIRED!) 🚨**

Before you can use the headless script or schedule the bot to run automatically, you **MUST** perform a manual login using the standard `ms_rewards_bot.py` script. This creates and populates the persistent browser profile that all subsequent runs will use to stay logged in.

1.  Run the standard script from your terminal:
    ```bash
    python ms_rewards_bot.py
    ```
2.  An Edge browser window will open. The script is configured to use a persistent profile (saved by default in `~/.ms_rewards_automation_profile` or `%USERPROFILE%\.ms_rewards_automation_profile`). It will navigate to the Microsoft account login page and wait for you.
3.  **Manually log in to your Microsoft account** in the browser window that appears. Enter your email/phone, password, and complete any multi-factor authentication steps if prompted.
4.  Once you successfully log in, the browser should redirect to your account or the Microsoft Rewards dashboard. The script will detect that it is no longer on a login page and will then proceed to run the daily tasks for the first time.
5.  After the tasks are completed, the browser window will close. The standard script (`ms_rewards_bot.py`) will then enter its internal scheduling loop. You can safely press `Ctrl+C` in your terminal to exit the script at this point, as your login session is now saved in the profile.

Your login credentials and session information are now saved within the persistent browser profile directory. All future runs (using either script) will use this saved profile and should log in automatically.

---

### Subsequent Runs (Visible - `ms_rewards_bot.py`)

After the initial login, you can run the standard script again anytime. It will use the saved profile and should log in automatically before performing tasks.

```bash
python ms_rewards_bot.py [arguments]
```

By default, after running tasks once, this script will then check for the next day's scheduled time (default 10:00 AM) and wait in a loop.

### Subsequent Runs (Headless - `ms_rewards_bot_headless.py`)

You can manually run the headless script to test it. It will use the saved profile, run tasks in the background without opening a window, and exit upon completion.

```bash
python ms_rewards_bot_headless.py [arguments]
```

This is the script you will typically use when setting up the bot as a background service.

### Command Line Arguments

Both `ms_rewards_bot.py` and `ms_rewards_bot_headless.py` accept arguments:

*   `--nosearch`: Add this flag to skip the Bing search tasks (both desktop and mobile). The script will still attempt to complete daily sets and other activities.
    ```bash
    python ms_rewards_bot.py --nosearch
    python ms_rewards_bot_headless.py --nosearch
    ```
*   `--time HH:MM`: **(Only used by `ms_rewards_bot.py` for its internal scheduler)** Specifies the daily time (in 24-hour `HH:MM` format) when the standard script should wake up and run tasks from its internal scheduling loop. Default is `10:00`. This argument is **ignored** by `ms_rewards_bot_headless.py` because scheduling is handled externally by system services.
    ```bash
    python ms_rewards_bot.py --time 08:00
    ```

### Setting up as a Background Service for Daily Runs

For automating the bot to run every day without needing to open a terminal, it's best to use your operating system's native tools to schedule the `ms_rewards_bot_headless.py` script.

This ensures reliability, proper logging, and that the bot runs even after a computer restart (depending on configuration).

Follow the detailed guide for your operating system in the `docs` folder:

*   [**Setting up as a Background Service (Linux & Windows)**](docs/background_service_setup.md)

## Logging

The script logs its activity to a file named `ms_rewards_automation.log`. This file is created in the directory from which the script is executed.

The log includes timestamps, actions taken (starting searches, completing activities), success/failure indications, and points balance checks. Check this file if you encounter issues or want to see the bot's progress.

When running as a background service, the location of this log file depends on the 'Working Directory' specified in your service or task configuration (e.g., your bot's root directory).

## Troubleshooting

*   **Login Timeout:** If the visible browser opens for login but times out, ensure you are actively logging into your Microsoft account within the 5-minute wait period. You must complete the full login process manually in the browser window that appears.
*   **Browser Not Opening / WebDriver Error:** Ensure Microsoft Edge is installed. The `webdriver-manager` library should automatically download the correct driver, but internet restrictions, firewalls, or insufficient user permissions can sometimes cause issues. Check the console output or `ms_rewards_automation.log` for specific error messages.
*   **Element Not Found / Script Fails Mid-Run:** Websites can change. If the script fails repeatedly at a specific step (e.g., finding the search box, clicking an activity card), the website's structure might have changed, and the script's element locators (like XPaths) may need updating. Check the log for `NoSuchElementException`, `TimeoutException`, etc.
*   **Background Service Issues (Windows/Linux):** Refer to the `docs/background_service_setup.md` guide. Common issues include incorrect file paths, insufficient permissions for the user running the task/service, or the service not starting correctly (e.g., needing a graphical session or user login).
*   **Persistent Profile Location:** The persistent profile is saved by default in a hidden folder in your user's home directory (`.ms_rewards_automation_profile`). Do not delete this folder after your first login, as it contains your session data.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
