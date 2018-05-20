using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ptlk_OPC;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ptlk_OPC.Tests
{
    [TestClass()]
    public class OPCTests
    {
        #region OPC
        [TestMethod()]
        public void GetTreeTest()
        {
            IOPC OPC = new OPC
            {
                ProgID = "ICONICS.ModbusOPC.3",
                Node = "127.0.0.1"
            };
            OPC.Connect();

            string tree = OPC.GetTree();
            if (tree == "[]")
            {
                Assert.Fail();
            }

            OPC.Disconnect();
        }

        [TestMethod()]
        public void SetGetValueTest()
        {
            IOPC OPC = new OPC
            {
                ProgID = "ICONICS.ModbusOPC.3",
                Node = "127.0.0.1"
            };
            OPC.Connect();

            string writeValue = "0";
            OPC.SetValue("T.A", writeValue);

            Thread.Sleep(500);

            string readValue = OPC.GetValue("T.A");
            if (readValue != writeValue)
            {
                Assert.Fail();
            }

            OPC.Disconnect();
        }

        [TestMethod()]
        public void SetGetGroupValueTest()
        {
            IOPC OPC = new OPC
            {
                ProgID = "ICONICS.ModbusOPC.3",
                Node = "127.0.0.1"
            };
            OPC.Connect();

            string[] items = new string[] { "T.A", "T.A1", "T.A2" };
            string[] resetValue = new string[] { "0", "0", "0" };
            string[] writeValue = new string[] { "1", "2", "3" };

            OPC.SetGroupItemID(ref items);
            OPC.SetGroupValue(ref resetValue);
            OPC.SetGroupValue(ref writeValue);

            Thread.Sleep(50);

            string[] readValue = OPC.GetGroupValue();

            for (int i = 0; i < readValue.Length; i++)
            {
                if (readValue[i] != writeValue[i].ToString())
                {
                    Assert.Fail();
                }
            }

            OPC.Disconnect();
        }
        #endregion

        #region OPC_XML
        [TestMethod()]
        public void GetTree_XMLTest()
        {
            IOPC OPC = new OPC
            {
                ProgID = "OPC_XML_DA_WrapperService.asmx",
                Node = "http://127.0.0.1/vdir"
            };
            OPC.Connect();

            string tree = OPC.GetTree();
            if (tree == "[]")
            {
                Assert.Fail();
            }

            OPC.Disconnect();
        }

        [TestMethod()]
        public void SetGetValue_XMLTest()
        {
            IOPC OPC = new OPC
            {
                ProgID = "OPC_XML_DA_WrapperService.asmx",
                Node = "http://127.0.0.1/vdir"
            };
            OPC.Connect();

            string writeValue = "0";
            OPC.SetValue("T.A", writeValue);

            Thread.Sleep(500);

            string readValue = OPC.GetValue("T.A");
            if (readValue != writeValue)
            {
                Assert.Fail();
            }

            OPC.Disconnect();
        }

        [TestMethod()]
        public void SetGetGroupValue_XMLTest()
        {
            IOPC OPC = new OPC
            {
                ProgID = "OPC_XML_DA_WrapperService.asmx",
                Node = "http://127.0.0.1/vdir"
            };
            OPC.Connect();

            string[] items = new string[] { "T.A", "T.A1", "T.A2" };
            string[] resetValue = new string[] { "0", "0", "0" };
            string[] writeValue = new string[] { "1", "2", "3" };

            OPC.SetGroupItemID(ref items);
            OPC.SetGroupValue(ref resetValue);
            OPC.SetGroupValue(ref writeValue);

            Thread.Sleep(50);

            string[] readValue = OPC.GetGroupValue();

            for (int i = 0; i < readValue.Length; i++)
            {
                if (readValue[i] != writeValue[i].ToString())
                {
                    Assert.Fail();
                }
            }

            OPC.Disconnect();
        }
        #endregion
    }
}