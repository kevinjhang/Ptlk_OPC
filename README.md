# Ptlk_OPC - OPC DA Read/Write Tools
Bridge Pattern by OPC DA Automation Wrapper 2.02|OPCDAAuto.dll|x86)

Support .NET and VB6 Environment

## Properties
### string ProgID
OPC ProgID ex: ICONICS.ModbusOPC.3(DCOM) or OPC_XML_DA_WrapperService.asmx(XML)
### string Node
OPC Node ex: 127.0.0.1(DCOM) or http://127.0.0.1/vdir(XML)
### int UpdateRate
Monitor update rate
### int PingTimeout
Check connection use
### int ConnectRate
Keep connection use
### bool IsConnected
Connection status

## Functions (All are sync)
### void Connect()
Connect OPC Server
### void GetTree()
Get OPC Tree(json)
### string GetValue(string ItemID)
Get Value
### void SetValue(string ItemID, string Value)
Set Value
### void SetGroupItemID(string[] ItemIDs)
Set Group ItemID
### string[] GetGroupValue()
Get Group Value
### void SetGroupValue(string[] Values)
Set Group Value
### void SetMonitorItemID(string[] ItemIDs)
Set Monitor ItemID
### void StartMonitor()
Start Monitor
### void StopMonitor()
Stop Monitor
### void Disconnect()
Disconnect OPC Server

## Events
### DataChange(int NumItems, object ClientHandles, object ItemValues, object TimeStamps)
Monitor callback
### EventLog(string Message)
Log callback
