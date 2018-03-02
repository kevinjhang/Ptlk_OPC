using System.Runtime.InteropServices;

namespace Ptlk_OPC
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOPC
    {
        event DataChangeHandler DataChange;
        event EventLogHandler EventLog;

        [DispId(0)]
        string ProgID { get; set; }
        [DispId(1)]
        string Node { get; set; }
        [DispId(2)]
        int UpdateRate { get; set; }
        [DispId(3)]
        int PingTimeout { get; set; }
        [DispId(4)]
        int ConnectRate { get; set; }
        [DispId(5)]
        bool IsConnected { get; }
        [DispId(6)]
        void Connect();
        [DispId(7)]
        string GetTree();
        [DispId(8)]
        string GetValue(string ItemID);
        [DispId(9)]
        void SetValue(string ItemID, string Value);
        [DispId(10)]
        void SetGroupItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs);
        [DispId(11)]
        string[] GetGroupValue();
        [DispId(12)]
        void SetGroupValue([MarshalAs(UnmanagedType.SafeArray)] ref string[] Values);
        [DispId(13)]
        void SetMonitorItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs);
        [DispId(14)]
        void StartMonitor();
        [DispId(15)]
        void StopMonitor();
        [DispId(16)]
        void Disconnect();
    }

    [ComVisible(false)]
    public delegate void DataChangeHandler(int NumItems, object ClientHandles, object ItemValues, object TimeStamps);

    [ComVisible(false)]
    public delegate void EventLogHandler(string Message);

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOPCEvents
    {
        [DispId(0)]
        void DataChange(int NumItems, object ClientHandles, object ItemValues, object TimeStamps);
        [DispId(1)]
        void EventLog(string Message);
    }
}
