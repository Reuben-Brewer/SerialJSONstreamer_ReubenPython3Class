###########################

SerialJSONstreamer_ReubenPython3Class

Control class (including ability to hook to Tkinter GUI) to read JSON strings sent over USB-Serial devices.

Reuben Brewer, Ph.D.

reuben.brewer@gmail.com

www.reubotics.com

Apache 2 License

Software Revision C, 12/26/2025

Verified working on:

Python 3.12/13

Windows 11 64-bit

Raspberry Pi Bookworm

(may work on Mac in non-GUI mode, but haven't tested yet)

Notes:

1. In Windows, you can get each sensor's USB-serial-device serial number by following the instructions in the USBserialDevice_GettingSerialNumberInWindows.png screenshot in this folder.

2. In Windows, you can manually set the latency timer for each sensor by following the instructions in the USBserialDevice_SettingLatencyTimerManuallyInWindows.png screenshot in this folder.

3. For serial messages to be parsed correctly, they must be in this format:
JSONstring CRChighByte CRClowByte \r\n

Example:
{"Time":735.252,"AnalogSensor_Value_AO0":0,"AngleFloatDegrees":193.9644754}CRCCRC\r\n

###########################

########################### Python module installation instructions, all OS's

############

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies: ['ftd2xx', 'LowPassFilterForDictsOfLists_ReubenPython2and3Class', 'ReubenGithubCodeModulePaths', 'serial', 'serial.tools']

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies_TestProgram: ['CSVdataLogger_ReubenPython3Class', 'EntryListWithBlinking_ReubenPython2and3Class', 'keyboard', 'LowPassFilterForDictsOfLists_ReubenPython2and3Class', 'MyPlotterPureTkinterStandAloneProcess_ReubenPython2and3Class', 'MyPrint_ReubenPython2and3Class', 'ReubenGithubCodeModulePaths']

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies_NestedLayers: ['EntryListWithBlinking_ReubenPython2and3Class', 'GetCPUandMemoryUsageOfProcessByPID_ReubenPython3Class', 'numpy', 'pexpect', 'psutil', 'pyautogui', 'ReubenGithubCodeModulePaths']

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies_All:['CSVdataLogger_ReubenPython3Class', 'EntryListWithBlinking_ReubenPython2and3Class', 'ftd2xx', 'GetCPUandMemoryUsageOfProcessByPID_ReubenPython3Class', 'keyboard', 'LowPassFilterForDictsOfLists_ReubenPython2and3Class', 'MyPlotterPureTkinterStandAloneProcess_ReubenPython2and3Class', 'MyPrint_ReubenPython2and3Class', 'numpy', 'pexpect', 'psutil', 'pyautogui', 'ReubenGithubCodeModulePaths', 'serial', 'serial.tools']

pip install psutil

pip install pyserial (NOT pip install serial).

pip install ftd2xx

Unzip FTDI_ D2XX_drivers__CDM2123620_Setup.zip, run the EXE, and then manually copy ftd2xx64.dll --> C:\Anaconda3\

############

############

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies: ['pandas', 'win32com.client', 'xlsxwriter', 'xlutils.copy', 'xlwt']

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies_TestProgram: []

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies_NestedLayers: []

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies_All:['pandas', 'win32com.client', 'xlsxwriter', 'xlutils.copy', 'xlwt']

pip install pywin32=311

pip install xlsxwriter==3.2.9 #Might have to manually delete older version from /lib/site-packages if it was distutils-managed. Works overall, but the function ".set_size" doesn't do anything.

pip install xlutils==2.0.0

pip install xlwt==1.3.0

############

###########################
