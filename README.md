# Ptlk_OPC - OPC DA Read/Write Tools

## Propertys
### string ProgID
### string Node
### int UpdateRate 
### int PingTimeout
### int ConnectRate
### bool IsConnected

## Functions
### void Connect()
### void GetTree()
### string GetValue(string ItemID)
### void SetValue(string ItemID, string Value)
### void SetGroupItemID(string[] ItemIDs)
### string[] GetGroupValue()
### void SetGroupValue(string[] Values)
### void SetMonitorItemID(string[] ItemIDs)
### void StartMonitor()
### void StopMonitor()
### void Disconnect()

## Events
### DataChange(int NumItems, object ClientHandles, object ItemValues, object TimeStamps)
### EventLog(string Message)
