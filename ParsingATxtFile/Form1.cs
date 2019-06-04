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

            //Initialize parse button as disabled.
            btnDasButton.Enabled = false;
        }





        //Choose file button --
        //On click opens a file explorer window for the user to navigate to the desired txt file.
        private void Button1_Click(object sender, EventArgs e)
        {
            //Open file explorer - Only show .txt files.
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            //Get filepath and verify .txt extension.
            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = choofdlog.FileName;

                lblFilePath.Text = sFileName;

                int indexaroni = lblFilePath.Text.IndexOf(".");
                string fileTypeVerification = lblFilePath.Text.Substring(indexaroni+1, 3);

                //Verify file type extension and enable Parse File button.
                if (fileTypeVerification == "txt")
                {
                    btnDasButton.Enabled = true;
                }
                else
                {
                    //error message for non-.txt files.
                    btnDasButton.Enabled = false;
                    MessageBox.Show("Please choose a file with a '.txt' extension.");
                }
            }
        }

        //Parse data from .txt file button --
        //Open an Excel document, export data from .txt (by whitespace), and fill cells of Excel doc by whitespace.
        private void btnDasButton_Click(object sender, EventArgs e)
        {
            //Init Lists of data from .txt file.
            List<string> theGoods = new List<string>();
            List<string> lines = File.ReadAllLines(lblFilePath.Text).ToList();

            //Open and init Excel doc.
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

                //Generic line/row 1 spacer
                oSheet.Cells[1, 1] = "TextFileLine#";
                oSheet.Cells[1, 2] = "LineTextData";

                oSheet.get_Range("A1", "C1").Font.Bold = true;
                oSheet.get_Range("A1", "C1").VerticalAlignment = excel.XlVAlign.xlVAlignCenter;

                for (int i = 1; i <= lines.Count; i++)
                {
                    //Parses data from the .txt file by line
                    List<string> entries = lines[i].Split(' ').ToList();
                    foreach (var entry in entries)
                    {
                        //Parses data from lines by whitespace
                        List<string> singleEntry = entry.Split(null).ToList();

                        //Inserts each data piece by cell
                        for (int o = 0; o < singleEntry.Count; o++)
                        {
                            if (singleEntry[o].StartsWith("0") || singleEntry[o].StartsWith("-") || singleEntry[o].StartsWith("1"))
                            {
                                oSheet.Cells[o + 2][i + 2] = singleEntry[o];
                            }
                        }
                    }

                    //This is specific to my company documents that include this in each line to space out a description
                    //Parses description of each line
                    int index01 = lines[i].IndexOf("Y\"");
                    if (index01 > 0)
                    {
                        //Puts description in Substring of the entire line
                        lines[i] = lines[i].Substring(0, index01 + 1);
                    }

                    //Numbers each line and inserts description data per line
                    oSheet.Cells[1][i + 1] = i;
                    oSheet.Cells[2][i + 1] = lines[i - 1];

                    //This adds the last line of data (line number and description) and ends the loop
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

