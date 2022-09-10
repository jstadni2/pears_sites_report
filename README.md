# PEARS Sites Report

The PEARS Sites Report compiles the site records created in [PEARS](https://www.k-state.edu/oeie/pears/) by [Illinois Extension](https://extension.illinois.edu/) staff on a monthly basis.
In order to prevent site duplication, select staff are authorized to manage requests for new site records. Other users are notified when they enter sites into PEARS without permission.

## Installation

The recommended way to install the PEARS Sites Report script is through git, which can be downloaded [here](https://git-scm.com/downloads). Once downloaded, run the following command:

```bash
git clone https://github.com/jstadni2/pears_sites_report
```

Alternatively, this repository can be downloaded as a zip file via this link:
[https://github.com/jstadni2/pears_sites_report/zipball/master/](https://github.com/jstadni2/pears_sites_report/zipball/master/)

This repository is designed to run out of the box on a Windows PC using Docker and the [/example_inputs](https://github.com/jstadni2/pears_sites_report/tree/master/example_inputs) and [/example_outputs](https://github.com/jstadni2/pears_sites_report/tree/master/example_outputs) directories.
To run the script in its current configuration, follow [this link](https://docs.docker.com/desktop/windows/install/) to install Docker Desktop for Windows. 

With Docker Desktop installed, this script can be run simply by double-clicking the `run_script.bat` file in your local directory.

The `run_script.bat` file can also be run in Command Prompt by entering the following command with the appropriate path:

```bash
C:\path\to\pears_sites_report\run_script.bat
```

### Setup instructions for SNAP-Ed implementing agencies

The following steps are required to execute the PEARS Sites Report script using your organization's PEARS data:
1. Contact [PEARS support](mailto:support@pears.io) to set up an [AWS S3](https://aws.amazon.com/s3/) bucket to store automated PEARS exports.
2. Download the automated PEARS exports. Illinois Extension's method for downloading exports from the S3 is detailed in the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting/blob/6f370389776fb8f88495fbe4e7918c203fd84997/pears_nightly_export_reformatting.py#L9-L45).
3. Set the appropriate input and output paths in `pears_sites_report.py` and `run_script.bat`.
	- The [Input Files](#input-files) and [Output Files](#output-files) sections provide an overview of required and output data files.
	- Copying input files to the build context would enable continued use of Docker and `run_script.bat` with minimal modifications.
4. Set the username and password variables in [pears_sites_report.py](https://github.com/jstadni2/pears_sites_report/blob/master/pears_sites_report.py#L93-L94) using valid Office 365 credentials.	

### Additional setup considerations

- The formatting of PEARS export workbooks changes periodically. The example PEARS exports included in the [/example_inputs](https://github.com/jstadni2/pears_sites_report/tree/master/example_inputs) directory are based on workbooks downloaded on 08/22/22.
Modifications to `pears_sites_report.py` may be necessary to run with subsequent PEARS exports.
- Illinois Extension utilized [Task Scheduler](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page) to run this script from a Windows PC on a monthly basis.
- Plans to deploy the PEARS Sites Report script on AWS were never implemented and are currently beyond the scope of this repository.
- Other SNAP-Ed implementing agencies intending to utilize the PEARS Sites Report script should consider the following adjustments as they pertain to their organization:
	- The `send_mail()` function in [pears_sites_report.py](https://github.com/jstadni2/pears_sites_report/blob/master/pears_sites_report.py#L139) is defined using Office 365 as the host.
	Change the host to the appropriate email service provider if necessary.
	
## Input Files

The following input files are required to run the PEARS Sites Report script:
- PEARS module exports:
    - [Site_Export.xlsx](https://github.com/jstadni2/pears_sites_report/blob/master/example_inputs/Site_Export.xlsx)
    - [User_Export.xlsx](https://github.com/jstadni2/pears_sites_report/blob/master/example_inputs/User_Export.xlsx)

Example input files are provided in the [/example_inputs](https://github.com/jstadni2/pears_sites_report/tree/master/example_inputs) directory. 
PEARS module exports included as example files are generated using the [Faker](https://faker.readthedocs.io/en/master/) Python package and do not represent actual program evaluation data. 

## Output Files

The following output files are produced by the PEARS Sites Report script:
- [PEARS Sites Report YYYY-MM.xlsx](https://github.com/jstadni2/pears_sites_report/blob/master/example_outputs/PEARS%20Sites%20Report%202022-07.xlsx): A workbook that compiles the site records entered into PEARS during the given month.

Example output files are provided in the [/example_outputs](https://github.com/jstadni2/pears_sites_report/tree/master/example_outputs) directory.
