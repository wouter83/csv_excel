using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace csv_excel
{
    public partial class Form1 : Form
    {
        public class ExcelCell
        {
            public string value;
            public string pos;
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.ShowDialog();
            string filename = dialog.FileName;
            Dictionary<int, List<ExcelCell>> DList = new Dictionary<int, List<ExcelCell>>();
            int i = 0;

            System.IO.StreamReader file = new System.IO.StreamReader(filename);
            // read the first line and parse the headers
            string headerLine = file.ReadLine();
           

            // Put the headers in a list so when we fill excel it will be placed correctly
            List<ExcelCell> headerlist = SplitString(headerLine);
            DList.Add(i++, headerlist);
            string line;
            // read a line, parse it, put it in a list.
            while ((line = file.ReadLine()) != null)
            {
                List<ExcelCell> rowList = SplitString(line);
                DList.Add(i++, rowList);
            }

            DisplayInExcel(DList);
        }

        private static List<ExcelCell> SplitString(string split)
        {
            // (["'])(?:(?=(\\?))\2.)*?\1
            string pattern = @"([""'])(?:(?=(\\?))\2.)*?\1";
            Regex rx = new Regex(pattern);
            MatchCollection matches = rx.Matches(split);

            List <ExcelCell> list = new List<ExcelCell>();
            for (int i = 0; i < matches.Count; ++i)
            {
                ExcelCell ec = new ExcelCell()
                {
                    pos = ((char)(0x41 + i)).ToString(),
                    value = matches[i].Value
                };
                list.Add(ec);
            }
            return list;
        }

        static void DisplayInExcel(Dictionary<int, List<ExcelCell>> DList)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            int columnCount = 0;
            // Establish column headings in cells A1 and B1.
            foreach(KeyValuePair<int, List<ExcelCell>> pair in DList)
            {
                columnCount = Math.Max(pair.Value.Count, columnCount);
                foreach(ExcelCell ec in pair.Value)
                {
                    workSheet.Cells[pair.Key +1, ec.pos.ToString()] = ec.value.TrimStart('"').TrimEnd('"');
                }
            }

            for (int i = 0; i < columnCount; ++i)
            {
                workSheet.Columns[i + 1].AutoFit();
            }
            for (int i = 0; i < columnCount; ++i)
            {
                ((Excel.Range)workSheet.Columns[i+1]).AutoFit();
            }
        }
    }
}
