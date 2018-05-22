using OPCAutomation;
using System;
using System.Collections.Generic;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Ptlk_OPC
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IOPCEvents))]
    [Guid("32030719-7F74-3953-A80D-5D70A16464BB")]
    public class OPC : IOPC, IDisposable
    {
        public event DataChangeHandler DataChange
        {
            add
            {
                if (m_OPC != null)
                {
                    m_OPC.DataChange += value;
                }
                m_DataChange.Add(value);
            }
            remove
            {
                if (m_OPC != null)
                {
                    m_OPC.DataChange -= value;
                }
                m_DataChange.Remove(value);
            }
        }

        public event EventLogHandler EventLog
        {
            add
            {
                if (m_OPC != null)
                {
                    m_OPC.EventLog += value;
                }
                m_EventLog.Add(value);
            }
            remove
            {
                if (m_OPC != null)
                {
                    m_OPC.EventLog -= value;
                }
                m_EventLog.Remove(value);
            }
        }

        public string ProgID { get; set; }
        public string Node { get; set; }

        public int UpdateRate
        {
            get
            {
                if (m_OPC != null)
                {
                    return m_OPC.UpdateRate;
                }
                return m_UpdateRate;
            }
            set
            {
                if (m_OPC != null)
                {
                    m_OPC.UpdateRate = value;
                }
                m_UpdateRate = value;
            }
        }

        public int PingTimeout
        {
            get
            {
                if (m_OPC != null)
                {
                    return m_OPC.PingTimeout;
                }
                return m_PingTimeout;
            }
            set
            {
                if (m_OPC != null)
                {
                    m_OPC.PingTimeout = value;
                }
                m_PingTimeout = value;
            }
        }

        public int ConnectRate
        {
            get
            {
                if (m_OPC != null)
                {
                    return m_OPC.ConnectRate;
                }
                return m_ConnectRate;
            }
            set
            {
                if (m_OPC != null)
                {
                    m_OPC.ConnectRate = value;
                }
                m_ConnectRate = value;
            }
        }

        public bool IsConnected
        {
            get
            {
                if (m_OPC != null)
                {
                    return m_OPC.IsConnected;

                }
                return false;
            }
        }

        public void SetGroupItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            m_OPC?.SetGroupItemID(ref ItemIDs);
            m_GroupItemID = ItemIDs;
        }

        public void SetMonitorItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            m_OPC?.SetMonitorItemID(ref ItemIDs);
            m_MonitorItemID = ItemIDs;
        }

        public void StartMonitor()
        {
            m_OPC?.StartMonitor();
            m_IsMonitor = true;
        }

        public void StopMonitor()
        {
            m_OPC?.StopMonitor();
            m_IsMonitor = false;
        }

        public OPC()
        {
        }

        public void Connect()
        {
            if (ProgID == null) throw new ArgumentNullException(nameof(ProgID));
            if (Node == null) throw new ArgumentNullException(nameof(Node));

            m_OPC?.Disconnect();

            if (ProgID.Contains("XML"))
            {
                m_OPC = new OPC_XML();
            }
            else
            {
                m_OPC = new OPC_DA();
            }

            foreach (var d in m_DataChange)
            {
                m_OPC.DataChange += d;
            }

            foreach (var d in m_EventLog)
            {
                m_OPC.EventLog += d;
            }

            m_OPC.ProgID = ProgID;
            m_OPC.Node = Node;
            m_OPC.UpdateRate = m_UpdateRate;
            m_OPC.PingTimeout = m_PingTimeout;
            m_OPC.ConnectRate = m_ConnectRate;
            m_OPC.SetGroupItemID(ref m_GroupItemID);
            m_OPC.SetMonitorItemID(ref m_MonitorItemID);
            if (m_IsMonitor) m_OPC.StartMonitor();
            m_OPC.Connect();
        }

        public string GetTree()
        {
            return m_OPC?.GetTree();
        }

        public string GetValue(string ItemID)
        {
            return m_OPC?.GetValue(ItemID);
        }

        public void SetValue(string ItemID, string Value)
        {
            m_OPC?.SetValue(ItemID, Value);
        }

        public string[] GetGroupValue()
        {
            return m_OPC?.GetGroupValue();
        }

        public void SetGroupValue([MarshalAs(UnmanagedType.SafeArray)] ref string[] Values)
        {
            m_OPC?.SetGroupValue(ref Values);
        }

        public void Disconnect()
        {
            m_OPC?.Disconnect();
            m_OPC = null;
            m_DataChange.Clear();
            m_EventLog.Clear();
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

        ~OPC()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        private IOPC m_OPC;
        private List<DataChangeHandler> m_DataChange;
        private List<EventLogHandler> m_EventLog;
        private string[] m_GroupItemID;
        private string[] m_MonitorItemID;
        private bool m_IsMonitor;
        private int m_UpdateRate = 1000;
        private int m_PingTimeout = 5000;
        private int m_ConnectRate = 30000;
    }
}
