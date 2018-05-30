using System;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Ptlk_OPC.OPC_XML_DA;

namespace Ptlk_OPC
{
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IOPCEvents))]
    [Guid("6D9F27BF-5E78-4BF9-8F96-6F6B975715FC")]
    class OPC_XML : IOPC, IDisposable
    {
        public event DataChangeHandler DataChange;
        public event EventLogHandler EventLog;

        private OPC_XML_DA_WrapperService OPCServer1;
        private string ServerHandlesM;
        private string[] GroupItemID;
        private string[] MonitorItemID;
        private bool IsMonitor;
        private DateTime ServerStartTime;
        private Timer Timer;
        private Timer Timer2;

        public string ProgID { get; set; } // OPC_XML_DA_WrapperService.asmx
        public string Node { get; set; }   // http://127.0.0.1/vdir
        public int UpdateRate { get; set; }
        public int PingTimeout { get; set; }
        public int ConnectRate { get; set; }
        public bool IsConnected { get; private set; }

        public OPC_XML()
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
                    string[] branches = null;
                    result = GetTreeItemByBranches(ref branches);
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
            string result = "*";

            if (string.IsNullOrEmpty(ItemID)) return result;

            try
            {
                if (!IsConnected) return result;

                if (OPCServer1 != null)
                {
                    var options = new RequestOptions();

                    var item = new ReadRequestItem
                    {
                        ItemPath = "",
                        ItemName = ItemID
                    };

                    var itemList = new ReadRequestItemList
                    {
                        Items = new[] { item }
                    };

                    OPCServer1.Read(options, itemList, out var replyList, out _);
                    var value = replyList.Items[0].Value ?? "*";
                    result = value.ToString();
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
            if (string.IsNullOrEmpty(ItemID)) return;

            try
            {
                if (!IsConnected) return;

                if (OPCServer1 != null)
                {
                    var options = new RequestOptions();

                    var item = new ItemValue
                    {
                        ItemPath = "",
                        ItemName = ItemID,
                        Value = Value
                    };

                    var itemList = new WriteRequestItemList
                    {
                        Items = new[] { item }
                    };

                    OPCServer1.Write(options, itemList, false, out _, out _);
                }
            }
            catch (Exception ex)
            {
                CheckConnected();
                Log($"{nameof(SetValue)}: {ex.Message} {nameof(ItemID)}: {ItemID} {nameof(Value)}: {Value}{Environment.NewLine}{ex.StackTrace}");
            }
        }

        public void SetGroupItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            if (ItemIDs == null) return;

            GroupItemID = ItemIDs;
        }

        public string[] GetGroupValue()
        {
            if (GroupItemID == null) return null;

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
                    var options = new RequestOptions();

                    var items = new ReadRequestItem[GroupItemID.Length];
                    for (int i = 0; i <= items.Length - 1; i++)
                    {
                        var item = new ReadRequestItem
                        {
                            ItemPath = "",
                            ItemName = GroupItemID[i]
                        };
                        items[i] = item;
                    }

                    var itemList = new ReadRequestItemList
                    {
                        Items = items
                    };

                    OPCServer1.Read(options, itemList, out var replyList, out _);
                    for (int i = 0; i <= result.Length - 1; i++)
                    {
                        result[i] = (replyList.Items[i].Value ?? "*").ToString();
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
                    var options = new RequestOptions();

                    var items = new ItemValue[GroupItemID.Length];
                    for (int i = 0; i <= items.Length - 1; i++)
                    {
                        var item = new ItemValue
                        {
                            ItemPath = "",
                            ItemName = GroupItemID[i],
                            Value = Values[i]
                        };
                        items[i] = item;
                    }

                    var itemList = new WriteRequestItemList
                    {
                        Items = items
                    };

                    OPCServer1.Write(options, itemList, false, out _, out _);
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
        }

        public void StartMonitor()
        {
            try
            {
                IsMonitor = true;

                if (!IsConnected) return;

                if (OPCServer1 != null)
                {
                    var options = new RequestOptions
                    {
                        ReturnItemName = true,
                        ReturnItemTime = true
                    };

                    var items = new SubscribeRequestItem[MonitorItemID.Length];
                    for (int i = 0; i <= items.Length - 1; i++)
                    {
                        var item = new SubscribeRequestItem
                        {
                            ItemPath = "",
                            ItemName = MonitorItemID[i]
                        };
                        items[i] = item;
                    }

                    var itemList = new SubscribeRequestItemList
                    {
                        Items = items
                    };

                    OPCServer1.Subscribe(options, itemList, true, int.MaxValue, out var replyList, out _, out ServerHandlesM);

                    Publish(replyList);

                    if (Timer2 == null)
                    {
                        Timer2 = new Timer(Timer2Callback, null, 0, Timeout.Infinite);
                    }
                    else
                    {
                        Timer2.Change(0, Timeout.Infinite);
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
                Timer2?.Dispose();
                Timer2 = null;

                IsMonitor = false;

                if (!IsConnected) return;

                if (OPCServer1 != null && ServerHandlesM != null)
                {
                    string _ = null;
                    OPCServer1.SubscriptionCancel(ServerHandlesM, ref _);
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

                OPCServer1?.Dispose();
                OPCServer1 = null;
                Log($"{nameof(Disconnect)}ed: {Node} {ProgID}");
                IsConnected = false;
            }
            catch (Exception ex)
            {
                Log($"{nameof(Disconnect)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
        }

        private void _Connect()
        {
            try
            {
                CheckConnected();
                if (!IsConnected)
                {
                    OPCServer1?.Dispose();
                    OPCServer1 = new OPC_XML_DA_WrapperService
                    {
                        Url = $"{Node}/{ProgID}"
                    };
                    ServicePointManager.Expect100Continue = false;
                    OPCServer1.GetStatus("", "", out var status);
                    ServerStartTime = status.StartTime;
                    Log($"{nameof(Connect)}ed: {Node} {ProgID}");
                    IsConnected = true;
                    if (IsMonitor) StartMonitor();
                }
            }
            catch (Exception ex)
            {
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
                if (OPCServer1 != null)
                {
                    if (ServerStartTime != null)
                    {
                        OPCServer1.GetStatus("", "", out var status);

                        if (ServerStartTime == status.StartTime)
                        {
                            result = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"{nameof(CheckConnected)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
            IsConnected = result;
        }

        private void Timer2Callback(object state)
        {
            try
            {
                if (OPCServer1 != null && ServerHandlesM != null)
                {
                    var options = new RequestOptions
                    {
                        ReturnItemName = true,
                        ReturnItemTime = true
                    };

                    OPCServer1.SubscriptionPolledRefresh(options, new[] { ServerHandlesM }, DateTime.MinValue
                        , false, 0, false, out string[] invalidHandles, out var replyList, out _, out _);

                    Publish(replyList);
                }
            }
            catch (Exception ex)
            {
                Log($"{nameof(Timer2Callback)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
            finally
            {
                Timer2?.Change(UpdateRate, Timeout.Infinite);
            }
        }

        private string GetTreeItemByBranches(ref string[] branches)
        {
            StringBuilder result = new StringBuilder();
            try
            {
                result.Append("[");
                string cp = null;
                BrowseElement[] element;

                if (branches?[0] == null)
                {
                    OPCServer1.Browse(null, null, null, "", "", ref cp, int.MaxValue, browseFilter.branch
                        , null, null, false, false, false, out element, out _, out _);
                }
                else
                {
                    OPCServer1.Browse(null, null, null, "", string.Join(".", branches), ref cp, int.MaxValue, browseFilter.branch
                        , null, null, false, false, false, out element, out _, out _);
                }
                int brancheCount = element.Length;
                string[] branches2 = new string[0];
                for (int i = 0; i < brancheCount; i++)
                {
                    if (branches?[0] == null)
                    {
                        OPCServer1.Browse(null, null, null, "", "", ref cp, int.MaxValue, browseFilter.branch
                            , null, null, false, false, false, out element, out _, out _);
                        branches2 = new string[1];
                        branches2[0] = element[i].Name;
                    }
                    else
                    {
                        OPCServer1.Browse(null, null, null, "", string.Join(".", branches), ref cp, int.MaxValue, browseFilter.branch
                            , null, null, false, false, false, out element, out _, out _);
                        Array.Resize(ref branches2, branches.Length + 1);
                        for (int j = 0; j < branches.Length; j++)
                        {
                            branches2[j] = branches[j];
                        }
                        branches2[branches2.Length - 1] = element[i].Name;
                    }
                    result.Append("{\"Name\":\"");
                    result.Append(element[i].Name);
                    result.Append("\",\"BrancheArray\":");
                    result.Append(GetTreeItemByBranches(ref branches2));
                    result.Append(",\"LeafArray\":[");
                    OPCServer1.Browse(null, null, null, "", string.Join(".", branches2), ref cp, int.MaxValue, browseFilter.item
                        , null, null, false, false, true, out element, out _, out _);
                    for (int j = 0; j < element.Length; j++)
                    {
                        result.Append("{\"Name\":\"");
                        result.Append(element[j].Name);
                        result.Append("\"}");
                        if (j != element.Length) result.Append(",");
                    }
                    result.Append("]}");
                    if (i != brancheCount) result.Append(",");
                }

                result.Append("]");
            }
            catch (Exception ex)
            {
                Log($"{nameof(GetTreeItemByBranches)}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
                result.Clear();
                result.Append("[]");
            }
            return result.ToString();
        }

        private void Publish(dynamic replyList)
        {
            ItemValue[] replyItems = new ItemValue[0];

            if (replyList is SubscribeReplyItemList subscribelist)
            {
                replyItems = new ItemValue[subscribelist.Items.Length];
                for (int i = 0; i < replyItems.Length; i++)
                {
                    replyItems[i] = subscribelist.Items[i].ItemValue;
                }
            }
            else if (replyList is SubscribePolledRefreshReplyItemList[] polledList)
            {
                if (polledList.Length == 0) return;
                replyItems = polledList[0].Items;
            }

            if (replyItems.Length > 0)
            {
                var clientHandles = new object[replyItems.Length];
                var itemValue = new object[replyItems.Length];
                var timeStamps = new object[replyItems.Length];

                for (int i = 0; i < replyItems.Length; i++)
                {
                    for (int j = 0; j < MonitorItemID.Length; j++)
                    {
                        if (MonitorItemID[j] == replyItems[i].ItemName)
                        {
                            clientHandles[i] = j;
                            break;
                        }
                    }
                    itemValue[i] = replyItems[i].Value;
                    timeStamps[i] = replyItems[i].Timestamp;
                }

                DataChange?.Invoke(replyItems.Length, CVarArrBase1(clientHandles), CVarArrBase1(itemValue), CVarArrBase1(timeStamps));
            }
        }

        private void Log(string Message)
        {
            EventLog?.Invoke(Message);
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

        ~OPC_XML()
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
