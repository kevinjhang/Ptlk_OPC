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
            OPC.SetValue("DO16.DO_0", writeValue);

            Thread.Sleep(500);

            string readValue = OPC.GetValue("DO16.DO_0");
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
            string[] s1 = new string[] { "DO16.DO_0", "DO16.DO_1", "DO16.DO_2" };
            OPC.SetGroupItemID(ref s1);

            string[] s2 = new string[] { "8", "8", "8" };
            OPC.SetGroupValue(ref s2);

            s1 = new string[] { "DO16.DO_0", "DO16.DO_1", "DO16.DO_2" };
            OPC.SetGroupItemID(ref s1);

            s2 = new string[] { "0", "1", "2" };
            OPC.SetGroupValue(ref s2);

            Thread.Sleep(500);

            string[] readValue = OPC.GetGroupValue();

            for (int i = 0; i < readValue.Length; i++)
            {
                if (readValue[i] != i.ToString())
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
                ProgID = "OPC_XML_DA_WrapperService",
                Node = "http://127.0.0.1/vdir/OPC_XML_DA_WrapperService.asmx"
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
                ProgID = "OPC_XML_DA_WrapperService",
                Node = "http://127.0.0.1/vdir/OPC_XML_DA_WrapperService.asmx"
            };
            OPC.Connect();
            string writeValue = "0";
            OPC.SetValue("DO16.DO_0", writeValue);

            Thread.Sleep(500);

            string readValue = OPC.GetValue("DO16.DO_0");
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
                ProgID = "OPC_XML_DA_WrapperService",
                Node = "http://127.0.0.1/vdir/OPC_XML_DA_WrapperService.asmx"
            };
            OPC.Connect();
            string[] s1 = new string[] { "DO16.DO_0", "DO16.DO_1", "DO16.DO_2" };
            OPC.SetGroupItemID(ref s1);

            string[] s2 = new string[] { "8", "8", "8" };
            OPC.SetGroupValue(ref s2);

            s1 = new string[] { "DO16.DO_0", "DO16.DO_1", "DO16.DO_2" };
            OPC.SetGroupItemID(ref s1);

            s2 = new string[] { "0", "1", "2" };
            OPC.SetGroupValue(ref s2);

            Thread.Sleep(500);

            string[] readValue = OPC.GetGroupValue();

            for (int i = 0; i < readValue.Length; i++)
            {
                if (readValue[i] != i.ToString())
                {
                    Assert.Fail();
                }
            }

            OPC.Disconnect();
        }
        #endregion

        #region OPC_XML Subscribe
        [TestMethod()]
        public void Subscribe_XMLTest()
        {
            var formTests = new OPC_XMLTests();
            formTests.Show();
            Application.Run(formTests);
        }
        #endregion
    }
}