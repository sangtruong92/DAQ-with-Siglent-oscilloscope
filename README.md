"# DAQ-with-Siglent-oscilloscope" 

This code use python code to acquire data from siglent oscilloscope (version: 1104X-E and 1204X-E). 
Data is saved .xlsx files (excel file)

To perform code:
1) Installing libraries from command line(Window), terminal(Ubuntu) before you run this code:
  + python get-pip.py 
  + pip install pyvisa  
  + pip install pyvisa pyvisa-py 
  + pip install pandas
    
2) Check IP adress of oscilloscope step by step:
  + Function of oscilloscope: Utility (change page) -> I/O -> IP Set.
  Note:
  + Make sure your computer connected to oscilloscope.
  + You can check the IP addess on web brower: http://192.168.1.212/ , for instance.
    
3) Change the same IP Adress of oscilloscope in code : REMOTE_IP = "169.254.144.94"
   

