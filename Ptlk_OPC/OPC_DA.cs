using OPCAutomation;
using System;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Ptlk_OPC
{
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IOPCEvents))]
    [Guid("F19DA3AA-D4A3-4BE9-BECA-BC1BCA5ACCA2")]
    class OPC_DA : IOPC, IDisposable
    {
        public event DataChangeHandler DataChange;
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
        private bool IsPingSuccess;
        private Timer Timer;

        public string ProgID { get; set; } // ICONICS.ModbusOPC.3
        public string Node { get; set; }   // 127.0.0.1
        public int UpdateRate { get; set; }
        public int PingTimeout { get; set; }
        public int ConnectRate { get; set; }
        public bool IsConnected { get; private set; }

        public OPC_DA()
        {
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
                    if (OPCBrowser1 != null)
                    {
                        OPCBrowser1.AccessRights = 0;
                        string[] branches = null;
                        result = GetTreeItemByBranches(ref branches);
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log($"{nameof(GetTree)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }

            return result;
        }

        public string GetValue(string ItemID)
        {
            string result = null;

            if (string.IsNullOrEmpty(ItemID)) return result;

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
                            if (OPCItem1 == null)
                            {
                                OPCItem1 = OPCGroup1.OPCItems.AddItem(ItemID, 0);
                            } 
                        }
                        finally
                        {
                            if (OPCItem1 != null)
                            {
                                OPCItem1.Read((short)OPCDataSource.OPCCache, out object value, out object quality, out object timeStamp);

                                if (!IsWaitingQuality(quality))
                                {
                                    if (!IsBadQuality(quality))
                                    {
                                        result = value.ToString();
                                    }
                                    else
                                    {
                                        result = null;
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
                Log($"{nameof(GetValue)}: {ex.Message} {nameof(ItemID)}: {ItemID}{Environment.NewLine}{ex.StackTrace}");
            }
            return result;
        }

        public void SetValue(string ItemID, string Value)
        {
            if (string.IsNullOrEmpty(ItemID) || string.IsNullOrEmpty(Value)) return;

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
                Log($"{nameof(SetValue)}: {ex.Message} {nameof(ItemID)}: {ItemID} {nameof(Value)}: {Value}{Environment.NewLine}{ex.StackTrace}");
                throw;
            }
        }

        public void SetGroupItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            if (ItemIDs == null) return;

            GroupItemID = ItemIDs;

            IsChangeGroupItemID = true;
        }

        public string[] GetGroupValue()
        {
            if (GroupItemID == null) return null;

            string[] result = new string[GroupItemID.Length];

            try
            {
                for (int i = 0; i <= result.Length - 1; i++)
                {
                    result[i] = null;
                }

                if (!IsConnected) return result;

                if (OPCServer1 != null)
                {
                    if (OPCGroupG != null)
                    {
                        if (IsChangeGroupItemID)
                        {
                            if (RemoveItemID(OPCGroupG, ServerHandlesG))
                            {
                                IsChangeGroupItemID = false;
                            }
                        }
                        if (AddItemID(OPCGroupG, GroupItemID, ref ServerHandlesG))
                        {
                            OPCGroupG.SyncRead(
                                (short)OPCDataSource.OPCCache,
                                ServerHandlesG.Length,
                                CIntArrBase1(ServerHandlesG),
                                out Array values,
                                out Array errors,
                                out object qualities,
                                out object timeStamps);

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
                                                result[i] = null;
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
                Log($"{nameof(GetGroupValue)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
            return result;
        }

        public void SetGroupValue([MarshalAs(UnmanagedType.SafeArray)] ref string[] Values)
        {
            if (Values == null) return;

            if (Values.Length != GroupItemID.Length) return;

            try
            {
                if (!IsConnected) return;

                if (OPCServer1 != null)
                {
                    if (OPCGroupG != null)
                    {
                        if (IsChangeGroupItemID)
                        {
                            if (RemoveItemID(OPCGroupG, ServerHandlesG))
                            {
                                IsChangeGroupItemID = false;
                            }
                        }
                        if (AddItemID(OPCGroupG, GroupItemID, ref ServerHandlesG))
                        {
                            OPCGroupG.SyncWrite(
                                ServerHandlesG.Length,
                                CIntArrBase1(ServerHandlesG),
                                CVarArrBase1(Values),
                                out Array Errors);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log($"{nameof(SetGroupValue)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
        }

        public void SetMonitorItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            if (ItemIDs == null) return;

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
                            if (RemoveItemID(OPCGroupM, ServerHandlesM))
                            {
                                IsChangeMonitorItemID = false;
                            }
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
                Log($"{nameof(StartMonitor)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
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
                Log($"{nameof(StopMonitor)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
        }

        public void Disconnect()
        {
            try
            {
                Timer?.Dispose();
                Timer = null;

                CheckConnected();
                if (IsPingSuccess && IsConnected)
                {
                    if (OPCServer1 != null)
                    {
                        OPCServer1.OPCGroups.RemoveAll();
                        OPCServer1.Disconnect();
                        Log($"{nameof(Disconnect)}ed: {Node} {ProgID}");
                    }
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
                Log($"{nameof(Disconnect)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
        }

        private IPStatus PingNode()
        {
            Ping pingSender = new Ping();
            PingOptions options = new PingOptions
            {
                DontFragment = true
            };
            // Create a buffer of 32 bytes of data to be transmitted.
            string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(data);
            PingReply reply = pingSender.Send(Node, PingTimeout, buffer, options);
            return reply.Status;
        }

        private void _Connect()
        {
            try
            {
                CheckConnected();
                if (IsPingSuccess && !IsConnected)
                {
                    OPCServer1 = new OPCServer();
                    OPCServer1.Connect(ProgID, Node);
                    Log($"{nameof(Connect)}ed: {Node} {ProgID}");
                    OPCGroup1 = OPCServer1.OPCGroups.Add();
                    OPCGroupG = OPCServer1.OPCGroups.Add();
                    OPCGroupM = OPCServer1.OPCGroups.Add();
                    OPCGroupM.DataChange += OPCGroupM_DataChange;
                    IsConnected = true;
                    if (IsMonitor) StartMonitor();
                }
            }
            catch (Exception ex)
            {
                OPCServer1 = null;
                Log($"{nameof(Connect)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
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
                    IsPingSuccess = true;
                    if (OPCServer1 != null)
                    {
                        int serverState = OPCServer1.ServerState;
                        if (serverState == (int)OPCServerState.OPCRunning)
                        {
                            result = true;
                        }
                    }
                }
                else
                {
                    IsPingSuccess = false;
                    OPCServer1 = null;
                    Log($"{nameof(PingNode)}: {Node} {ping}");
                }
            }
            catch (PingException pex)
            {
                IsPingSuccess = false;
                OPCServer1 = null;
                Log($"{nameof(CheckConnected)}: {pex.Message} {nameof(Node)}: {Node} {nameof(PingTimeout)}{PingTimeout}");
            }
            catch (ExternalException eex)
            {
                OPCServer1 = null;
                Log($"{nameof(CheckConnected)}: {eex.Message} {Environment.NewLine}{eex.StackTrace}");
            }
            catch (Exception ex)
            {
                Log($"{nameof(CheckConnected)}: {ex.Message} {Environment.NewLine}{ex.StackTrace}");
            }
            IsConnected = result;
        }

        private bool AddItemID(OPCGroup OPCGroup, string[] ItemID, ref int[] ServerHandles)
        {
            int i = 0;
            try
            {
                if (OPCGroup != null)
                {
                    if (OPCGroup.OPCItems.Count == 0)
                    {
                        Array.Resize(ref ServerHandles, ItemID.Length);

                        for (i = 0; i <= ServerHandles.Length - 1; i++)
                        {
                            ServerHandles[i] = OPCGroup.OPCItems.AddItem(ItemID[i], i).ServerHandle;
                        }
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log($"{nameof(AddItemID)}: {ex.Message} {nameof(ItemID)}: {ItemID}{Environment.NewLine}{ex.StackTrace}");
            }
            return false;
        }

        private bool RemoveItemID(OPCGroup OPCGroup, int[] ServerHandles)
        {
            int i = 0;
            try
            {
                if (OPCGroup != null)
                {
                    if (OPCGroup.OPCItems.Count > 0)
                    {
                        for (i = 0; i <= ServerHandles.Length - 1; i++)
                        {
                            Array ServerHandle = Array.CreateInstance(typeof(int), new int[] { 1 }, new int[] { 1 });
                            ServerHandle.SetValue(ServerHandles[i], 1);
                            if (((int)ServerHandle.GetValue(1)) != 0)
                            {
                                OPCGroup.OPCItems.Remove(1, ServerHandle, out Array Errors);
                            }
                        }
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log($"{nameof(RemoveItemID)}: {ex.Message} {Environment.NewLine}{ex.StackTrace}");
            }
            return false;
        }

        private bool IsBadQuality(object Quality)
        {
            bool result;
            result = false;
            if (Quality != null)
            {
                if (int.TryParse(Quality.ToString(), out int quality))
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
                if (int.TryParse(Quality.ToString(), out int quality))
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
            try
            {
                result.Append("[");
                if (OPCBrowser1 != null)
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
                        result.Append("{\"Name\":\"");
                        result.Append(OPCBrowser1.Item(i));
                        result.Append("\",\"BrancheArray\":");
                        result.Append(GetTreeItemByBranches(ref branches2));
                        result.Append(",\"LeafArray\":[");
                        OPCBrowser1.MoveTo(CStrArrBase1(branches2));
                        OPCBrowser1.ShowLeafs();
                        for (int j = 1; j <= OPCBrowser1.Count; j++)
                        {
                            result.Append("{\"Name\":\"");
                            result.Append(OPCBrowser1.Item(j));
                            result.Append("\"}");
                            if (j != OPCBrowser1.Count) result.Append(",");
                        }
                        result.Append("]}");
                        if (i != brancheCount) result.Append(",");
                    }
                }
                result.Append("]");
            }
            catch (Exception ex)
            {
                Log($"{nameof(GetTreeItemByBranches)}: {ex.Message} {Environment.NewLine}{ex.StackTrace}");
                result.Clear();
                result.Append("[]");
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

        private Array CIntArrBase1(int[] Source)
        {
            Array result = Array.CreateInstance(typeof(int), new int[] { Source.Length }, new int[] { 1 });
            for (int i = 1; i <= Source.Length; i++)
            {
                result.SetValue(Source[i - 1], i);
            }
            return result;
        }

        private Array CVarArrBase1(object[] Source)
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
                    ItemValues.SetValue(null, i);
                }
            }
            DataChange?.Invoke(NumItems, ClientHandles, ItemValues, TimeStamps);
        }

        private void Log(string Message)
        {
            EventLog?.Invoke(Message);
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                Disconnect();
                disposedValue = true;
            }
        }

        ~OPC_DA()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
