# Task Runner

A simple graphical task scheduler for Windows that allows you to run scripts and programs on a schedule using cron-like syntax.

## Features

- Graphical user interface for easy task management
- Cron-like scheduling syntax
- Support for Python scripts and other executable files
- Log retention management
- Persistent task storage
- Task status monitoring

## Installation

1. Make sure you have Python 3.x installed on your system
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the task runner:
   ```
   python task_runner.py
   ```

2. To add a new task:
   - Click "Browse" to select the script or program you want to run
   - Enter the cron expression (format: minute hour day month weekday)
   - Set the log retention period in days
   - Click "Save Task"

3. Cron Expression Format:
   - minute (0-59)
   - hour (0-23)
   - day (1-31)
   - month (1-12)
   - weekday (0-6, where 0 is Monday)
   - Use * for any value

Example cron expressions:
- `* * * * *` - Run every minute
- `0 * * * *` - Run every hour
- `0 0 * * *` - Run once a day at midnight
- `0 0 * * 0` - Run once a week on Sunday at midnight

## Logs

Task execution logs are stored in `task_runner.log`. The log retention setting determines how long these logs are kept.

## Task Storage

Tasks are automatically saved to `tasks.json` and will be restored when you restart the application.
