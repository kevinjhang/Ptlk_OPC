using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OPCAutomation;
using System.Threading;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Ptlk_OPC
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOPC
    {
        [DispId(0)]
        string ProgID { get; set; }
        [DispId(1)]
        string Node { get; set; }
        [DispId(2)]
        int UpdateRate { get; set; }
        [DispId(3)]
        int ConnectTimeout { get; set; }
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

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOPCEvents
    {
        [DispId(0)]
        void DataChange(int NumItems, object ClientHandles, object ItemValues, object TimeStamps);
        [DispId(1)]
        void EventLog(string Message);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IOPCEvents))]
    [Guid("32030719-7F74-3953-A80D-5D70A16464BB")]
    public class OPC : IOPC, IDisposable
    {
        [ComVisible(false)]
        public delegate void DataChangeHandler(int NumItems, object ClientHandles, object ItemValues, object TimeStamps);
        public event DataChangeHandler DataChange;
        [ComVisible(false)]
        public delegate void EventLogHandler(string Message);
        public event EventLogHandler EventLog;

        private OPCServer OPCServer1;
        private OPCBrowser OPCBrowser1;
        private OPCGroup OPCGroup1;
        private OPCGroup OPCGroupG;
        private OPCGroup OPCGroupM;
        private int[] ServerHandlesG;
        private int[] ServerHandlesM;
        private string[] GroupItemID;
        private bool IsChangeGroupItemID;
        private string[] MonitorItemID;
        private bool IsChangeMonitorItemID;
        private bool IsMonitor;
        private Timer Timer;
        private string mProgID = "ICONICS.ModbusOPC.3";
        private string mNode = "127.0.0.1";
        private int mUpdateRate = 1000;
        private int mConnectTimeout = 3000;
        private int mConnectRate = 5000;
        private bool mIsConnected;
        private bool disposed;

        public string ProgID
        {
            get { return mProgID; }
            set { mProgID = value; }
        }

        public string Node
        {
            get { return mNode; }
            set { mNode = value; }
        }

        public int UpdateRate
        {
            get { return mUpdateRate; }
            set { mUpdateRate = value; }
        }

        public int ConnectRate
        {
            get { return mConnectRate; }
            set { mConnectRate = value; }
        }

        public int ConnectTimeout
        {
            get { return mConnectTimeout; }
            set { mConnectTimeout = value; }
        }

        public bool IsConnected
        {
            get { return mIsConnected; }
            private set { mIsConnected = value; }
        }

        public OPC()
        {
        }
        ~OPC()
        {
            Dispose(false);
        }

        public void Connect()
        {
            if (Timer == null)
            {
                _Connect();
                Timer = new Timer(TimerCallback, null, ConnectRate, Timeout.Infinite);
            }
            else
            {
                Timer.Change(0, Timeout.Infinite);
            }
        }

        public string GetTree()
        {
            string result = "[]";
            try
            {
                if (!IsConnected) return result;
                if (OPCServer1 != null)
                {
                    OPCBrowser1 = OPCServer1.CreateBrowser();
                    OPCBrowser1.AccessRights = 0;
                    string[] branches = null;
                    result = GetTreeItemByBranches(ref branches);
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("GetTreeItem:" + ex.Message + "\r\n" + ex.StackTrace);
            }
            return result;
        }

        public string GetValue(string ItemID)
        {
            string result = "*";
            try
            {
                if (!IsConnected) return result;
                if (OPCServer1 != null)
                {
                    if (OPCGroup1 != null)
                    {
                        OPCItem OPCItem1 = null;
                        try
                        {
                            OPCItem1 = OPCGroup1.OPCItems.Item(ItemID);
                        }
                        catch
                        {
                            if (OPCItem1 == null) OPCItem1 = OPCGroup1.OPCItems.AddItem(ItemID, 0);
                        }
                        finally
                        {
                            if (OPCItem1 != null)
                            {
                                object value, quality, timeStamp;
                                OPCItem1.Read((short)OPCDataSource.OPCCache, out value, out quality, out timeStamp);
                                if (!IsWaitingQuality(quality))
                                {
                                    if (!IsBadQuality(quality))
                                    {
                                        result = value.ToString();
                                    }
                                    else
                                    {
                                        result = "*";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("GetValue:" + ex.Message + " ItemID:" + ItemID + "\r\n" + ex.StackTrace);
            }
            return result;
        }

        public void SetValue(string ItemID, string Value)
        {
            try
            {
                if (!IsConnected) return;
                if (OPCServer1 != null)
                {
                    if (OPCGroup1 != null)
                    {
                        OPCItem OPCItem1 = null;
                        try
                        {
                            OPCItem1 = OPCGroup1.OPCItems.Item(ItemID);
                        }
                        catch
                        {
                            if (OPCItem1 == null) OPCItem1 = OPCGroup1.OPCItems.AddItem(ItemID, 0);
                        }
                        finally
                        {
                            OPCItem1?.Write(Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("SetValue:" + ex.Message + " ItemID:" + ItemID + "\r\n" + ex.StackTrace);
            }
        }

        public void SetGroupItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            GroupItemID = ItemIDs;
            IsChangeGroupItemID = true;
        }

        public string[] GetGroupValue()
        {
            string[] result = new string[GroupItemID.Length];
            try
            {
                for (int i = 0; i <= result.Length - 1; i++)
                {
                    result[i] = "*";
                }
                if (!IsConnected) return result;
                if (OPCServer1 != null)
                {
                    if (OPCGroupG != null)
                    {
                        if (IsChangeGroupItemID)
                        {
                            RemoveItemID(OPCGroupG, ServerHandlesG);
                            IsChangeGroupItemID = false;
                        }
                        if (AddItemID(OPCGroupG, GroupItemID, ref ServerHandlesG))
                        {
                            Array values, errors;
                            object qualities, timeStamps;
                            OPCGroupG.SyncRead((short)OPCDataSource.OPCCache, ServerHandlesG.Length, CLngArrBase1(ServerHandlesG), out values, out errors, out qualities, out timeStamps);
                            for (int i = 0; i <= result.Length - 1; i++)
                            {
                                object quality = ((Array)qualities).GetValue(i + 1);
                                if (values != null)
                                {
                                    if (values.GetValue(i + 1) != null)
                                    {
                                        if (!IsWaitingQuality(quality))
                                        {
                                            if (!IsBadQuality(quality))
                                            {
                                                result[i] = values.GetValue(i + 1).ToString();
                                            }
                                            else
                                            {
                                                result[i] = "*";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("GetGroupValue:" + ex.Message + "\r\n" + ex.StackTrace);
            }
            return result;
        }

        public void SetGroupValue([MarshalAs(UnmanagedType.SafeArray)] ref string[] Values)
        {
            try
            {
                if (!IsConnected) return;
                if (OPCServer1 != null)
                {
                    if (OPCGroupG != null)
                    {
                        if (IsChangeGroupItemID)
                        {
                            RemoveItemID(OPCGroupG, ServerHandlesG);
                            IsChangeGroupItemID = false;
                        }
                        if (AddItemID(OPCGroupG, GroupItemID, ref ServerHandlesG))
                        {
                            Array Errors;
                            OPCGroupG.SyncWrite(ServerHandlesG.Length, CLngArrBase1(ServerHandlesG), CVarArrBase1(Values), out Errors);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("SetGroupValue:" + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        public void SetMonitorItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            MonitorItemID = ItemIDs;
            IsChangeMonitorItemID = true;
        }

        public void StartMonitor()
        {
            try
            {
                IsMonitor = true;
                if (!IsConnected) return;
                if (OPCServer1 != null)
                {
                    if (OPCGroupM != null)
                    {
                        if (IsChangeMonitorItemID)
                        {
                            RemoveItemID(OPCGroupM, ServerHandlesM);
                            IsChangeMonitorItemID = false;
                        }
                        if (AddItemID(OPCGroupM, MonitorItemID, ref ServerHandlesM))
                        {
                            OPCGroupM.UpdateRate = UpdateRate;
                            OPCGroupM.IsActive = true;
                            OPCGroupM.IsSubscribed = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("StartMonitor:" + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        public void StopMonitor()
        {
            try
            {
                IsMonitor = false;
                if (!IsConnected) return;
                if (OPCServer1 != null)
                {
                    if (OPCGroupM != null)
                    {
                        OPCGroupM.IsSubscribed = false;
                        OPCGroupM.IsActive = false;
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log("StopMonitor:" + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        public void Disconnect()
        {
            try
            {
                Timer?.Dispose();
                Timer = null;
                CheckConnected();
                if (IsConnected)
                {
                    OPCServer1.OPCGroups.RemoveAll();
                    Log("Call Disconnect " + ProgID + " " + Node);
                    OPCServer1.Disconnect();
                    Log("Disconnected " + ProgID + " " + Node);
                }
                OPCServer1 = null;
                OPCBrowser1 = null;
                OPCGroup1 = null;
                OPCGroupG = null;
                OPCGroupM = null;
                IsConnected = false;
            }
            catch (Exception ex)
            {
                Log("Disconnect:" + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        private IPStatus PingNode()
        {
            Ping pingSender = new Ping();
            PingOptions options = new PingOptions();
            options.DontFragment = true;
            // Create a buffer of 32 bytes of data to be transmitted.
            string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(data);
            PingReply reply = pingSender.Send(Node, ConnectTimeout, buffer, options);
            return reply.Status;
        }

        private void _Connect()
        {
            try
            {
                CheckConnected();
                if (!IsConnected)
                {
                    OPCServer1 = new OPCServer();
                    Log("Call Connect " + ProgID + " " + Node);
                    OPCServer1.Connect(ProgID, Node);
                    Log("Connected " + ProgID + " " + Node);
                    if (OPCServer1 != null)
                    {
                        OPCGroup1 = OPCServer1.OPCGroups.Add();
                        OPCGroupG = OPCServer1.OPCGroups.Add();
                        OPCGroupM = OPCServer1.OPCGroups.Add();
                        OPCGroupM.DataChange += OPCGroupM_DataChange;
                        IsConnected = true;
                        if (IsMonitor) StartMonitor();
                    }
                }
            }
            catch (Exception ex)
            {
                Log("_Connect:" + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        private void TimerCallback(object state)
        {
            _Connect();
            Timer?.Change(ConnectRate, Timeout.Infinite);
        }

        private void CheckConnected()
        {
            bool result = false;
            try
            {
                IPStatus ping = PingNode();
                if (ping == IPStatus.Success)
                {
                    if (OPCServer1 != null)
                    {
                        int serverState = OPCServer1.ServerState;
                        if (serverState == (int)OPCServerState.OPCRunning)
                        {
                            result = true;
                        }
                        else
                        {
                            Log("ServerState:" + GetStateString(serverState));
                        }
                    }
                }
                else
                {
                    Log("Ping:" + ping);
                }
            }
            catch (PingException pex)
            {
                Log("Ping:" + pex.Message);
            }
            catch (ExternalException eex)
            {
                OPCServer1 = null;
                Log("External:" + eex.Message);
            }
            catch (Exception ex)
            {
                Log("CheckConnected:" + ex.Message + "\r\n" + ex.StackTrace);
            }
            IsConnected = result;
        }

        private bool AddItemID(OPCGroup OPCGroup, string[] ItemID, ref int[] ServerHandles)
        {
            try
            {
                if (OPCGroup != null)
                {
                    if (OPCGroup.OPCItems.Count == 0)
                    {
                        Array.Resize(ref ServerHandles, ItemID.Length);

                        for (int i = 0; i <= ServerHandles.Length - 1; i++)
                        {
                            ServerHandles[i] = OPCGroup.OPCItems.AddItem(ItemID[i], i).ServerHandle;
                        }
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log("AddItemID:" + ex.Message + "\r\n" + ex.StackTrace);
            }
            return false;
        }

        private void RemoveItemID(OPCGroup OPCGroup, int[] ServerHandles)
        {
            try
            {
                if (OPCGroup != null)
                {
                    if (OPCGroup.OPCItems.Count > 0)
                    {
                        Array Errors;
                        for (int i = 0; i <= ServerHandles.Length - 1; i++)
                        {
                            Array ServerHandle = Array.CreateInstance(typeof(int), new int[] { 1 }, new int[] { 1 });
                            ServerHandle.SetValue(ServerHandles[i], 1);
                            if (((int)ServerHandle.GetValue(1)) != 0)
                            {
                                OPCGroup.OPCItems.Remove(1, ServerHandle, out Errors);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("RemoveItemID:" + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        private bool IsBadQuality(object Quality)
        {
            bool result;
            result = false;
            if (Quality != null)
            {
                int quality;
                if (int.TryParse(Quality.ToString(), out quality))
                {
                    if (quality < 192)
                    {
                        result = true;
                    }
                }
            }
            return result;
        }

        private bool IsWaitingQuality(object Quality)
        {
            bool result;
            result = false;
            if (Quality != null)
            {
                int quality;
                if (int.TryParse(Quality.ToString(), out quality))
                {
                    if (quality == 32)
                    {
                        result = true;
                    }
                }
            }
            return result;
        }

        private string GetTreeItemByBranches(ref string[] branches)
        {
            StringBuilder result = new StringBuilder();
            result.Append("[");
            try
            {
                if (branches?[0] == null)
                {
                    OPCBrowser1.MoveToRoot();
                    OPCBrowser1.ShowBranches();
                }
                else
                {
                    OPCBrowser1.MoveTo(CStrArrBase1(branches));
                    OPCBrowser1.ShowBranches();
                }
                int brancheCount = OPCBrowser1.Count;
                string[] branches2 = new string[0];
                for (int i = 1; i <= brancheCount; i++)
                {
                    if (branches?[0] == null)
                    {
                        OPCBrowser1.MoveToRoot();
                        OPCBrowser1.ShowBranches();
                        branches2 = new string[1];
                        branches2[0] = OPCBrowser1.Item(i);
                    }
                    else
                    {
                        OPCBrowser1.MoveTo(CStrArrBase1(branches));
                        OPCBrowser1.ShowBranches();
                        Array.Resize(ref branches2, branches.Length + 1);
                        for (int j = 0; j <= branches.Length - 1; j++)
                        {
                            branches2[j] = branches[j];
                        }
                        branches2[branches2.Length - 1] = OPCBrowser1.Item(i);
                    }
                    result.Append("{\"Name\":\"" + OPCBrowser1.Item(i) + "\",\"BrancheArray\":" + GetTreeItemByBranches(ref branches2) + ",\"LeafArray\":[");
                    OPCBrowser1.MoveTo(CStrArrBase1(branches2));
                    OPCBrowser1.ShowLeafs();
                    for (int j = 1; j <= OPCBrowser1.Count; j++)
                    {
                        result.Append("{\"Name\":\"" + OPCBrowser1.Item(j) + "\"}");
                        if (j != OPCBrowser1.Count) result.Append(",");
                    }
                    result.Append("]}");
                    if (i != brancheCount) result.Append(",");
                }
                result.Append("]");
            }
            catch (Exception ex)
            {
                Log("GetTreeItemByBranches:" + ex.Message + "\r\n" + ex.StackTrace);
            }
            return result.ToString();
        }

        //Change Array Base, Because Array of OPCDAAuto.dll is Base 1
        private Array CStrArrBase1(string[] Source)
        {
            Array result = Array.CreateInstance(typeof(string), new int[] { Source.Length }, new int[] { 1 });
            for (int i = 1; i <= Source.Length; i++)
            {
                result.SetValue(Source[i - 1], i);
            }
            return result;
        }

        private Array CLngArrBase1(int[] Source)
        {
            Array result = Array.CreateInstance(typeof(int), new int[] { Source.Length }, new int[] { 1 });
            for (int i = 1; i <= Source.Length; i++)
            {
                result.SetValue(Source[i - 1], i);
            }
            return result;
        }

        private Array CVarArrBase1(string[] Source)
        {
            Array result = Array.CreateInstance(typeof(object), new int[] { Source.Length }, new int[] { 1 });
            for (int i = 1; i <= Source.Length; i++)
            {
                result.SetValue(Source[i - 1], i);
            }
            return result;
        }

        private string GetStateString(int State)
        {
            switch (State)
            {
                case 1:
                    return "Running";
                case 2:
                    return "Failed";
                case 3:
                    return "Noconfig";
                case 4:
                    return "Suspended";
                case 5:
                    return "Test";
                case 6:
                    return "Disconnected";
                default:
                    return "Null";
            }
        }

        //Monitor Callback
        private void OPCGroupM_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            for (int i = 1; i <= NumItems; i++)
            {
                if (IsBadQuality(Qualities.GetValue(i)))
                {
                    ItemValues.SetValue("*", i);
                }
            }
            DataChange?.Invoke(NumItems, ClientHandles, ItemValues, TimeStamps);
        }

        private void Log(string Message)
        {
            EventLog?.Invoke(Message);
        }

        private void Wait(Func<bool> predicate, long millisecondsTimeout)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            while (predicate() && millisecondsTimeout > sw.ElapsedMilliseconds)
            {
                Thread.Sleep(100);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                Disconnect();
                disposed = true;
            }
        }
    }
}
