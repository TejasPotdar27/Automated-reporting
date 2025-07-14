[README Test Summary Tool.docx](https://github.com/user-attachments/files/21192264/README.Test.Summary.Tool.docx)

Test Summary Tool - README
Overview
The Test Summary Tool is a powerful automation solution designed specifically for automotive testing teams using the Vector VT System. This tool transforms the time-consuming process of manually reviewing dozens or hundreds of test reports into an efficient, automated workflow that completes in minutes.
What once took a team hours can now be done by a single person in minutes.
Originally developed to address the challenge of processing 50-200 test reports generated during automated overnight and weekend testing cycles, the tool automatically scans Vector CANoe test reports, extracts essential data, and generates comprehensive Excel summaries with professional formatting.
Features
•	🚀 Batch Processing: Process hundreds of test reports simultaneously
•	📊 Automated Data Extraction: Extract verdicts (Pass/Fail/Inconclusive), durations, campaign details, and setup information from Vector CANoe test reports (.pdf, .vtestreport)
•	📈 Excel Report Generation: Generate ready-to-use, professionally styled Excel summaries
•	🔄 PDF Creation: Convert .vtestreport files to PDF format using Vector CANoe Test Report Viewer
•	🧹 Delivery Folder Cleanup: Remove unnecessary files while preserving essential logs and reports
•	💻 User-Friendly GUI: Simple Tkinter interface - no command-line experience required
•	⚙️ Customizable Verdict Keywords: Easily adapt to different testing environments and verdict schemes
Prerequisites
•	Windows OS (optimized for Windows environments)
•	Python 3.8+
•	Vector CANoe Test Report Viewer (required for PDF conversion functionality)
Installation
1.	Clone or download this repository
git clone [repository-url]
cd test-summary-tool

2.	Install required dependencies
pip install pandas openpyxl pdfplumber send2trash

3.	Verify Vector CANoe Test Report Viewer installation
o	The tool will automatically detect the installation
o	Ensure it's properly installed for PDF conversion features
Usage
Basic Usage
1.	Launch the tool
python sortnexcel_v3_9.py

2.	Configure settings in the GUI
o	Input Folder: Select the root directory containing your date-stamped TS_* test folders
o	Output Folder: Choose destination for Excel summary files
o	Verdict Keywords: (Optional) Customize keywords for flagging specific verdicts (default: Fail,Inconclusive)
3.	Execute operations
o	Generate Excel summary reports
o	Create or delete report PDFs
o	Clean up folders for delivery/archiving
Folder Structure
The tool expects the following folder structure:
Root Directory/
-> TS_YYYY-MM-DD_HH-MM-SS/   //Testcase1
    -> Report_test_report_1.vtestreport
    -> Report_test_report_1.pdf
    -> canLog.blf
-> TS_YYYY-MM-DD_HH-MM-SS/   //Testcase2
    -> Report_test_report_2.vtestreport
    -> Report_test_report_2.pdf
    -> canLog.blf


Key Benefits
Before	After
Manual Process: Hours of manual review by multiple team members	Automated Process: Minutes of automated processing by one person
Error-Prone: Manual data extraction and Excel entry	Reliable: Automated extraction with consistent formatting
Time-Intensive: 50-200 reports × manual review time	Efficient: Batch processing of all reports simultaneously
Resource-Heavy: Multiple team members required	Streamlined: Single operator workflow

Important Notes
•	Always review the generated Excel summary for completeness after processing
•	Use caution with cleanup/deletion features - verify folder selections before proceeding
•	Vector CANoe Test Report Viewer must be installed for PDF conversion functionality
•	The tool automatically detects and handles various Vector CANoe report formats
Contributing
We welcome contributions, feedback, and feature requests! Please feel free to:
•	Submit bug reports or feature requests via issues
•	Fork the repository and submit pull requests
•	Share your experience and suggestions for improvements
License
This project is licensed under the MIT License - see the LICENSE file for details.
Support
For questions, issues, or support:
•	Check the Issues section for known problems and solutions
•	Create a new issue for bug reports or feature requests
•	Review the documentation for troubleshooting tips
Transform your testing workflow today - from hours of manual work to minutes of automated efficiency with the Test Summary Tool.

  
