Scheduler
=========
This repo contains Python script and examples to build fair schedule for a list of people.

Input:
=====
The input is an XLSX file, containing 
-	List of people (optional: personal constraints)
-	Per position, a sheet with description, such as team size per hour, shift time etc.
-	Schedule for the previous day. We need this information to make sure that “time to rest” is observed, starting from the first hour.
	Note: this sheet must exist and have a predefined format, but the data cells can be empty.
	At Example.xlsx, there are two sheets that can serve as an input schedule:
		2023-11-02: empty
		2023-11-03: full

Please use format provided in Example.xlsx
It is recommended to see the demo

Output:
=======
Output is a schedule for the next N days, printed to screen and (optionally) back to the same XLS file
The following are observed while building the schedule:
-	Minimal time to rest after a shift:
	Default of 9 hours after night shift (configurable from command line)
	Default of 4 hours after day shift (configurable from command line)
-	A person doing a night shift, will not be assigned for another night shift the following night
-	The algorithm strives for maximal fairness. Fairness analysis is printed at the end of the run.

Run command:
=============

Usage: scheduler.py [-h] [--seed SEED] [--prev PREV] [--next NEXT] [--write] [--days DAYS] [--positions POSITIONS] [--ttrn TTRN] [--ttrd TTRD] XLS_file_name

Positional arguments:
  file_name             XLS file name

Required arguments:
  --prev <date>         Prev schedule sheet name (optional, default is today's date)
  --positions <N>       Number of positions
  --days <M>            Number of days to schedule

Optional arguments:
  -h, --help            show this help message and exit
  --seed SEED           Seed
  --next NEXT           Next schedule sheet name (optional, default is tomorrow's date)
  --write               Do write result to the XLS file
  --ttrn TTRN           Minimum time to rest after NIGHT shift
  --ttrd TTRD           Minimum time to rest after DAY shift

Feedback:
=========
Any feedback is welcome.
If you report a bug, please send the following:
-	Problem description
-	XLS file
-	Command line, better with seed

We will make an effort to address your requests, but no response time is guaranteed.
The project is 100% volunteering and we do it in our spare time
Hope you find it useful 😊

Lena Barzilai
Lena.barzilai@gmail.com
0544524290
