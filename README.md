# Power-System-Tools-for-PSSE
Python Codes for PSSE Software Maximization, Manipulation, Data Extraction, and etc.

# Bus Voltage Extraction Tool

## Overview
This Python script automates the extraction of bus voltage magnitudes (in p.u. or kV) from a power system simulation.  
It was developed to support **transmission planning studies** such as thermal, voltage, and short-circuit assessments.

## Features
- Connects to PSS®E simulation results (via `psspy`).
- Extracts bus voltage data for a given case.
- Saves results into Excel for easy reporting and analysis.

## Technologies Used
- **Python**
- **PSS®E API (psspy)**
- **excelpy** for Excel integration
