###########################

SerialJSONstreamer_ReubenPython3Class

Control class (including ability to hook to Tkinter GUI) to read JSON strings sent over USB-Serial devices.

Reuben Brewer, Ph.D.

reuben.brewer@gmail.com

www.reubotics.com

Apache 2 License

Software Revision A, 10/27/2024

Verified working on:

Python 3.12

Windows 11 64-bit

Raspberry Pi Buster

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

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies: ['ftd2xx', 'LowPassFilterForDictsOfLists_ReubenPython2and3Class', 'serial', 'serial.tools']

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies_TestProgram: ['CSVdataLogger_ReubenPython3Class', 'EntryListWithBlinking_ReubenPython2and3Class', 'keyboard', 'LowPassFilterForDictsOfLists_ReubenPython2and3Class', 'MyPlotterPureTkinterStandAloneProcess_ReubenPython2and3Class', 'MyPrint_ReubenPython2and3Class']

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies_NestedLayers: ['future.builtins', 'LowPassFilter_ReubenPython2and3Class', 'numpy', 'pexpect', 'psutil']

SerialJSONstreamer_ReubenPython3Class, ListOfModuleDependencies_All:['CSVdataLogger_ReubenPython3Class', 'EntryListWithBlinking_ReubenPython2and3Class', 'ftd2xx', 'future.builtins', 'keyboard', 'LowPassFilter_ReubenPython2and3Class', 'LowPassFilterForDictsOfLists_ReubenPython2and3Class', 'MyPlotterPureTkinterStandAloneProcess_ReubenPython2and3Class', 'MyPrint_ReubenPython2and3Class', 'numpy', 'pexpect', 'psutil', 'serial', 'serial.tools']

pip install psutil

pip install pyserial (NOT pip install serial).

############

############

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies: ['pandas', 'win32com.client', 'xlsxwriter', 'xlutils.copy', 'xlwt']

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies_TestProgram: []

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies_NestedLayers: []

ExcelPlot_CSVdataLogger_ReubenPython3Code_SerialJSONstreamer.py, ListOfModuleDependencies_All:['pandas', 'win32com.client', 'xlsxwriter', 'xlutils.copy', 'xlwt']

pip install pywin32         #version 305.1 as of 10/17/24

pip install xlsxwriter      #version 3.2.0 as of 10/17/24. Might have to manually delete older version from /lib/site-packages if it was distutils-managed. Works overall, but the function ".set_size" doesn't do anything.

pip install xlutils         #version 2.0.0 as of 10/17/24

pip install xlwt            #version 1.3.0 as of 10/17/24

############

###########################
