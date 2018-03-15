using OPCAutomation;
using System;
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
                if (m_OPC == null)
                {
                    m_DataChange += value;
                }
                else
                {
                    m_OPC.DataChange += value;
                }
            }
            remove
            {
                if (m_OPC == null)
                {
                    m_DataChange -= value;
                }
                else
                {
                    m_OPC.DataChange -= value;
                }
            }
        }

        public event EventLogHandler EventLog
        {
            add
            {
                if (m_OPC == null)
                {
                    m_EventLog += value;
                }
                else
                {
                    m_OPC.EventLog += value;
                }
            }
            remove
            {
                if (m_OPC == null)
                {
                    m_EventLog -= value;
                }
                else
                {
                    m_OPC.EventLog -= value;
                }
            }
        }

        public string ProgID { get => m_ProgID; set => m_ProgID = value; }
        public string Node { get => m_Node; set => m_Node = value; }

        public int UpdateRate
        {
            get
            {
                if (m_OPC == null)
                {
                    return m_UpdateRate;
                }
                else
                {
                    return m_OPC.UpdateRate;
                }
            }
            set
            {
                if (m_OPC == null)
                {
                    m_UpdateRate = value;
                }
                else
                {
                    m_OPC.UpdateRate = value;
                }
            }
        }

        public int PingTimeout
        {
            get
            {
                if (m_OPC == null)
                {
                    return m_PingTimeout;
                }
                else
                {
                    return m_OPC.PingTimeout;
                }
            }
            set
            {
                if (m_OPC == null)
                {
                    m_PingTimeout = value;
                }
                else
                {
                    m_OPC.PingTimeout = value;
                }
            }
        }

        public int ConnectRate
        {
            get
            {
                if (m_OPC == null)
                {
                    return m_ConnectRate;
                }
                else
                {
                    return m_OPC.ConnectRate;
                }
            }
            set
            {
                if (m_OPC == null)
                {
                    m_ConnectRate = value;
                }
                else
                {
                    m_OPC.ConnectRate = value;
                }
            }
        }

        public bool IsConnected
        {
            get
            {
                if (m_OPC == null)
                {
                    return false;
                }
                else
                {
                    return m_OPC.IsConnected;
                }
            }
        }

        public OPC()
        {
        }

        public void Connect()
        {
            if (ProgID == null || Node == null) return;

            if (m_OPC != null) m_OPC.Disconnect();

            if (ProgID.Contains("XML"))
            {
                m_OPC = new OPC_XML();
            }
            else
            {
                m_OPC = new OPC_DA();
            }

            m_OPC.ProgID = ProgID;
            m_OPC.Node = Node;
            m_OPC.UpdateRate = m_UpdateRate;
            m_OPC.PingTimeout = m_PingTimeout;
            m_OPC.ConnectRate = m_ConnectRate;
            m_OPC.DataChange += m_DataChange;
            m_OPC.EventLog += m_EventLog;
            m_OPC.SetGroupItemID(ref m_GroupItemID);
            m_OPC.SetMonitorItemID(ref m_MonitorItemID);
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

        public void SetGroupItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            if (m_OPC == null)
            {
                m_GroupItemID = ItemIDs;
            }
            else
            {
                m_OPC.SetGroupItemID(ref ItemIDs);
            }
        }

        public string[] GetGroupValue()
        {
            return m_OPC?.GetGroupValue();
        }

        public void SetGroupValue([MarshalAs(UnmanagedType.SafeArray)] ref string[] Values)
        {
            m_OPC?.SetGroupValue(ref Values);
        }

        public void SetMonitorItemID([MarshalAs(UnmanagedType.SafeArray)] ref string[] ItemIDs)
        {
            if (m_OPC == null)
            {
                m_MonitorItemID = ItemIDs;
            }
            else
            {
                m_OPC.SetMonitorItemID(ref ItemIDs);
            }
        }

        public void StartMonitor()
        {
            m_OPC?.StartMonitor();
        }

        public void StopMonitor()
        {
            m_OPC?.StopMonitor();
        }

        public void Disconnect()
        {
            m_OPC?.Disconnect();
            m_OPC = null;
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
        private event DataChangeHandler m_DataChange;
        private event EventLogHandler m_EventLog;
        private string[] m_GroupItemID;
        private string[] m_MonitorItemID;
        private string m_ProgID;
        private string m_Node;
        private int m_UpdateRate = 1000;
        private int m_PingTimeout = 5000;
        private int m_ConnectRate = 30000;
    }
}
