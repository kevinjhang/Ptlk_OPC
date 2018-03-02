using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ptlk_OPC.Tests
{
    public partial class OPC_XMLTests : Form
    {
        IOPC OPC;

        public OPC_XMLTests()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;

            OPC = new OPC
            {
                ProgID = "vdir/OPC_XML_DA_WrapperService",
                Node = "61.220.65.61"
            };

            OPC.Connect();
            string[] s1 = new string[] { "DO16.DO_0", "DO16.DO_1", "DO16.DO_2", "DO16.DO_3" };
            OPC.SetMonitorItemID(ref s1);
            OPC.DataChange += OPC_DataChange;
        }

        private void OPC_DataChange(int NumItems, object ClientHandles, object ItemValues, object TimeStamps)
        {
            for (int i = 0; i < NumItems; i++)
            {
                var handles = ((Array)ClientHandles).GetValue(i + 1);
                var tb = this.Controls.Find("textBox" + handles, false)[0] as TextBox;
                tb.Text = ((Array)ItemValues).GetValue(i + 1).ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OPC.StartMonitor();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OPC.StopMonitor();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OPC.Connect();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string s = OPC.GetTree();
            MessageBox.Show(s);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OPC.Disconnect();
        }
    }
}
