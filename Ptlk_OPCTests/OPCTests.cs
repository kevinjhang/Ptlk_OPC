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
        [TestMethod()]
        public void GetTreeItemTest()
        {
            OPC OPC = new OPC();
            OPC.ProgID = "ICONICS.ModbusOPC.3";
            OPC.Node = "127.0.0.1";
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
            OPC OPC = new OPC();
            OPC.ProgID = "ICONICS.ModbusOPC.3";
            OPC.Node = "127.0.0.1";
            OPC.Connect();
            string writeValue = DateTime.Now.Second.ToString();
            OPC.SetValue("OPC.Item1", writeValue);
            Thread.Sleep(1000);
            string readValue = OPC.GetValue("OPC.Item1");
            if (readValue != writeValue)
            {
                Assert.Fail();
            }
            OPC.Disconnect();
        }

        [TestMethod()]
        public void SetGetGroupValueTest()
        {
            OPC OPC = new OPC();
            OPC.ProgID = "ICONICS.ModbusOPC.3";
            OPC.Node = "127.0.0.1";
            OPC.Connect();
            string[] s1 = new string[] { "OPC.Item1" };
            OPC.SetGroupItemID(ref s1);
            string writeValue = DateTime.Now.Second.ToString();
            string[] s2 = new string[] { writeValue };
            OPC.SetGroupValue(ref s2);
            Thread.Sleep(1000);
            string readValue = OPC.GetGroupValue()[0];
            if (readValue != writeValue)
            {
                Assert.Fail();
            }
            OPC.Disconnect();
        }
    }
}