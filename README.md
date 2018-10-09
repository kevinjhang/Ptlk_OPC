# Ptlk_OPC - OPC DA Read/Write Tools
Facade by OPC DA Automation Wrapper 2.02 (OPCDAAuto.dll, x86)

Support .NET and VB6 Environment

## Properties
### string ProgID
OPC progID (ex: ICONICS.ModbusOPC.3, DCOM or OPC_XML_DA_WrapperService.asmx, XML)
### string Node
OPC node (ex: 127.0.0.1, DCOM or http://127.0.0.1/vdir, XML)
### int UpdateRate
Monitor update rate
### int PingTimeout
Use when check connection
### int ConnectRate
Use when keep connection
### bool IsConnected
Connection status

## Functions (All are sync)
### void Connect()
Connect OPC server
### void GetTree()
Get OPC tree(json)
### string GetValue(string ItemID)
Get value
### void SetValue(string ItemID, string Value)
Set value
### void SetGroupItemID(string[] ItemIDs)
Set group itemID
### string[] GetGroupValue()
Get group value
### void SetGroupValue(string[] Values)
Set group value
### void SetMonitorItemID(string[] ItemIDs)
Set monitor itemID
### void StartMonitor()
Start monitor
### void StopMonitor()
Stop monitor
### void Disconnect()
Disconnect OPC server

## Events
### DataChange(int NumItems, Array ClientHandles, Array ItemValues, Array TimeStamps)
Monitor callback
### EventLog(string Message)
Log callback
