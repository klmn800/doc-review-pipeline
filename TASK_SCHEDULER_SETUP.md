# Windows Task Scheduler Setup

This guide shows how to set up Windows Task Scheduler to run `watcher.py` on a recurring schedule. The watcher only invokes the pipeline when there are documents to process, so scheduling it frequently is cheap.

## Prerequisites

- `pipeline.py` and `watcher.py` are configured and you've successfully run the pipeline at least once manually
- `config.ini` exists with a valid Anthropic API key and an `input_folder` path
- The input folder is reachable from the user account that the task will run as (especially relevant for cloud-synced folders like OneDrive or Dropbox)

## Setup Instructions

### Step 1: Open Task Scheduler

1. Press `Windows + R` to open the Run dialog
2. Type `taskschd.msc` and press Enter

### Step 2: Create a New Task

1. In the right panel, click **Create Task...** (not "Create Basic Task")

### Step 3: General Tab

**Name:** `Document Review Watcher`

**Description:** `Runs the document review pipeline on a schedule. No-op when the inbox is empty.`

**Security options:**
- Check **Run whether user is logged on or not**
- Check **Run with highest privileges**
- Set **Configure for:** to your Windows version

### Step 4: Triggers Tab

Click **New...** to add one or more triggers. A common pattern is three runs per day at 9 AM, 1 PM, and 5 PM:

- **Begin the task:** On a schedule
- **Settings:** Daily
- **Start:** today's date at the desired time
- **Recur every:** 1 day
- **Enabled:** checked

Repeat for each additional time slot.

### Step 5: Actions Tab

Click **New...**

- **Action:** Start a program
- **Program/script:** `python` (or the full path to your Python executable, e.g. `C:\Python312\python.exe` — see Troubleshooting below)
- **Add arguments:** `watcher.py`
- **Start in:** the absolute path to your repo folder, e.g. `C:\Users\<you>\code\doc-review-pipeline`

### Step 6: Conditions Tab

**Power:**
- Check **Start the task only if the computer is on AC power** (recommended for laptops)
- Uncheck **Stop if the computer switches to battery power**
- Optionally check **Wake the computer to run this task**

**Network:** check **Start only if the following network connection is available** and select **Any connection** (the pipeline calls the Anthropic API).

### Step 7: Settings Tab

- Check **Allow task to be run on demand**
- Check **Run task as soon as possible after a scheduled start is missed**
- Set **If the task is already running:** to **Do not start a new instance**

### Step 8: Save

Click **OK**. Enter your Windows password when prompted (required so the task can run when you're not logged in).

### Step 9: Test

In Task Scheduler, find your task, right-click → **Run**. Then check `watcher.log` in the repo folder — you should see a recent entry. If there were documents in the inbox, you should also see a new entry in `tracker.xlsx`.

## Monitoring

- **Task Scheduler:** the **Last Run Time** and **Last Run Result** columns. `0x0` means success.
- **Watcher log:** `watcher.log` in the repo folder logs every invocation, including no-op runs.
- **Pipeline logs:** `logs/check_*.txt` for empty runs, `logs/review_*.txt` for runs that processed documents.

## Troubleshooting

**Task shows as "Running" but nothing happens**
Python isn't on the system PATH for the user the task runs as. Use the full path to your Python executable in the **Program/script** field — find it by running `where python` in a Command Prompt.

**Task fails immediately**
Check `watcher.log` in the repo folder. Common causes:
- `config.ini` not found — make sure the **Start in** path points to the repo folder, not the `python.exe` folder
- `input_folder` path doesn't exist or isn't readable as the task user
- API key is invalid

**Task doesn't run at scheduled time**
Verify triggers are enabled and that the **Conditions** tab requirements are met (e.g. AC power, network available). On laptops, AC-power-only is a common reason scheduled tasks don't fire.

## Modifying the Schedule

Right-click the task → **Properties** → **Triggers** tab → **Edit...** to change times, or **New...** to add additional runs per day.

## Disabling or Removing

- **Pause:** right-click → **Disable**
- **Resume:** right-click → **Enable**
- **Remove:** right-click → **Delete**
