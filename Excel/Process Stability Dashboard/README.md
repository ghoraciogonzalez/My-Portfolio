# Process Stability Analysis in Excel

**NOTE⚠️**: To use the scroll bar and macros included in this Excel workbook, it is necessary to download the file and enable both Excel functions and macros. Without enabling macros, the automated navigation and interactive features of the dashboard will not function properly.

This project aims to evaluate the stability of a bbatch measurement process using statistical tools implemented directly in Microsoft Excel. An interactive dashboard was developed to analyze process behavior through control charts, distribution analysis, and an automated navigation system powered by VBA macros.

# Project Targets
1. Quick identification of batch process stability.
2. Analyze the data distribution to assess whether the sample follows normality, enabling the application of process capability statistical studies.
3. Improved control and organization through a dynamic and automated dashboard.

# Analysis Methodology
The analysis was carried out in several stages:

  *Control Charts*
  1. Individual (I) and Moving Range (MR) charts were implemented to evaluate process stability across batches.

  *Normality Evaluation*
  1. **Histogram**: Used to visualize the shape of the distribution.
  2. **Q-Q Plot**: Compared the dataset against a theoretical normal distribution.
  3. **Measures of Central Tendency**: Mean, median, and mode were considered as preliminary indicators of distribution symmetry.

**NOTE**: This is a quick and preliminary analysis. More in-depth studies are required, including formal normality tests (e.g., Kolmogorov-Smirnov or Anderson-Darling) and process capability studies (Cp, Cpk, Ppk).

  *Interactive Dashboard*
  1. A selection bar (using Excel Developer tools) allows users to choose the batch for analysis.
  2. Each batch consists of 50 measurements.
  3. Dynamic ranges were configured to ensure proper dashboard functionality when navigating between batches.

# VBA Automation
VBA code snippets were included to enhance user interaction:
  1. **Tab Management**: The data sheet is automatically hidden when switching tabs.
  2. **Controlled Access**: From the dashboard, a macro-enabled button temporarily grants access to the hidden data sheet to add new measurements.

# Technologies Used
  1. Microsoft Excel: Dashboards, formulas, and charts.
  2. VBA (Visual Basic for Applications): Process automation.
