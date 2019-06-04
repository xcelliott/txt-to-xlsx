using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Text.RegularExpressions;



namespace ParsingATxtFile
{
    public partial class Form1 : Form
    {       
        public Form1()
        {
            InitializeComponent();

            btnDasButton.Enabled = false;
        }








        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = choofdlog.FileName;

                lblFilePath.Text = sFileName;

                int indexaroni = lblFilePath.Text.IndexOf(".");
                string fileTypeVerification = lblFilePath.Text.Substring(indexaroni+1, 3);

                if (fileTypeVerification == "txt")
                {
                    btnDasButton.Enabled = true;
                }
                else
                {
                    btnDasButton.Enabled = false;
                    MessageBox.Show("Please choose a file with a '.txt' extension.");
                }
            }
        }


        private void btnDasButton_Click(object sender, EventArgs e)
        {
            List<string> theGoods = new List<string>();
            List<string> lines = File.ReadAllLines(lblFilePath.Text).ToList();

            excel.Application oXL;
            excel.Workbook oWB;
            excel.Worksheet oSheet;
            excel.Range oRange;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                oXL = new excel.Application();
                oXL.Visible = true;

                oWB = (excel.Workbook)oXL.Workbooks.Add("");
                oSheet = (excel.Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 1] = "TextFileLine#";
                oSheet.Cells[1, 2] = "LineTextData";

                oSheet.get_Range("A1", "C1").Font.Bold = true;
                oSheet.get_Range("A1", "C1").VerticalAlignment = excel.XlVAlign.xlVAlignCenter;

                for (int i = 1; i <= lines.Count; i++)
                {

                    List<string> entries = lines[i].Split(' ').ToList();
                    foreach (var entry in entries)
                    {
                        List<string> singleEntry = entry.Split(null).ToList();
                        

                        for (int o = 0; o < singleEntry.Count; o++)
                        {
                            if (singleEntry[o].StartsWith("0") || singleEntry[o].StartsWith("-") || singleEntry[o].StartsWith("1"))
                            {
                                oSheet.Cells[o + 2][i + 2] = singleEntry[o];
                            }
                        }
                    }


                    int index01 = lines[i].IndexOf("Y\"");
                    if (index01 > 0)
                    {
                        lines[i] = lines[i].Substring(0, index01 + 1);
                    }

                    oSheet.Cells[1][i + 1] = i;
                    oSheet.Cells[2][i + 1] = lines[i - 1];

                    if (i == lines.Count - 1)
                    {
                        oSheet.Cells[1][lines.Count + 1] = lines.Count;
                        oSheet.Cells[2][lines.Count + 1] = lines[i];
                    }
                }

            }
            catch (Exception p)
            { Console.WriteLine("Exception: " + p); }
        }
    }
}

