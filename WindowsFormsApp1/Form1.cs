using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOIReadExcel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnReadExcel_Click(object sender, EventArgs e)
        {
            string errInfo = "";
            CNPOIReadExcel excel = new CNPOIReadExcel();

            if (!excel.OpenExcel(@"C:\Users\11468\Desktop\test.xlsx", ref errInfo)) {

                MessageBox.Show(errInfo);
                return;
            }

            string ret = excel.GetCell(1, 1);
            MessageBox.Show(ret);


            excel.CloseExcel();

        }
    }
}
