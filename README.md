# Computer Resource Monitor

A single Jupyter notebook, `monitor.ipynb`, that monitors a Windows machine. It samples
all computer resources along with every running process and Windows service, establishes a
benchmark of normal operating conditions, and detects and alarms when something deviates
from normal. Results are written to Excel files with a timestamp on every row of every
sheet.

This is a working version. Further improvements are planned.

## What it collects each sample

- CPU usage, overall and per core
- CPU temperature (where the sensor is readable)
- Memory: total, used, available, percentage
- Disk, per drive: total, used, free, percentage
- Battery: charge, plugged in or not, estimated time remaining
- Network, per interface: data sent and received since the last sample, plus totals
- GPU usage and memory
- System uptime and last boot time
- Every running process: id, name, CPU share, memory, status
- Every Windows service: name, display name, status, start type

## Benchmark and alarm

This is the core of the project. Three methods each flag deviations from normal, and a
concluding step combines them:

1. Fixed thresholds: plain limits from the config (CPU, memory, disk, GPU, temperature too
   high, battery too low, a required service stopped).
2. Statistical baseline: learns each metric's normal range from the early part of the run
   and flags later samples that fall outside it.
3. Saved baseline: saves the normal profile to a file so future runs compare against a
   reusable benchmark.

The concluding alarm merges all three, writes them to an alarms sheet in both Excel files,
and prints the overall verdict: all clear, or an alarm with a breakdown.

## Outputs

All in the `outputs` folder:

- `dashboard/dashboard.xlsx` - overwritten every sample, the latest snapshot only.
- `log/log.xlsx` - appended every sample, the full history of the run.
- `charts/` - PNG trend images built at the end of the run.
- `benchmark.json` - the saved normal profile for reusable comparison.

Both Excel files carry these sheets, all timestamped: system_metrics, disk, network,
processes, services, and alarms. The dashboard also gains trends and charts sheets at the
end of the run.

A shareable `monitor.html`, the exported notebook with all results, sits in the project
root for handing to anyone without Python.

## How to run

1. Open `monitor.ipynb` in VS Code or Jupyter.
2. Run the cells from the top: the config cell, then the dependencies setup cell (it
   installs anything missing), then the resource and benchmark sections.
3. Run `run_full_monitor()` to do a full run: it collects samples, raises the concluding
   alarm, and builds the charts. With no arguments it uses the config settings, every 5
   minutes for 24 hours.

For a quick demonstration instead of the full day, pass a short interval and duration:

    from datetime import timedelta
    run_full_monitor(interval=timedelta(seconds=10), duration=timedelta(seconds=30))

## Dependencies

- Python 3 (already installed; the notebook cannot install Python itself).
- `psutil`, `openpyxl`, `matplotlib` - checked and installed automatically by the
  dependencies setup cell, which also enforces a minimum `psutil` version needed for disk
  reads on Python 3.12.

## Adjusting the configuration

Everything adjustable lives in the config cell at the top: the sample interval, the total
run duration, the output paths, the alarm thresholds, the list of required services, and
the statistical baseline settings. Change them there and rerun.
