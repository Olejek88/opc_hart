Universal Hart OPC Server v0.90

installation notes:
-------------------
After install software you must register server with key /r. You also may unregistered server with key /u.
For example opc.exe /r or opc.exe /u.
Show help key /?.

version history:
----------------
v0.90 build 44
= fixed copy data from port to tag (while >2 zero in string early be 
interpreted as '\0' character);
= finally fixed timestamp on work;

v0.90 build 6
= fixed output date bug;
= fixed error output (kg/h) value;
= now show serial number correctly;
= fixed timestamp on connect;
- delete many records in log file;

v0.89 build 12
+ add ability configurate com-port throw ini file;
+ add evaluating and checking control sum;
+ loVendorInfo now show correct version;
+ add help key /?;
- delete many records in log file;

v0.88a
+ support 33 write and read tags;
+ support OPC DA Specifications 2.0;