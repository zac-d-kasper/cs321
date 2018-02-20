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

namespace Spreadsheet_KasperZac
{
    public partial class Form1 : Form
    {
        public SpreadsheetEngine.Spreadsheet mainSheet; 

        public Form1()
        {
            InitializeComponent();
            mainSheet = new SpreadsheetEngine.Spreadsheet(50, 26); 
            for (int i = 1; i <= 26; i++)
            {
                dataGridView1.Columns.Add(GetColFromIndex(i), GetColFromIndex(i));
            }
            for (int i = 1; i <= 50; i++)
            {
                dataGridView1.Rows.Add("");
                dataGridView1.Rows[i - 1].HeaderCell.Value = i.ToString();
            }
            mainSheet.CellChanged += OnCellChanged;

            /*mainSheet.CellGrid[0,0].CellText = "String";
            mainSheet.CellGrid[0,1].CellText = "=A1";
            mainSheet.CellGrid[0, 5].CellText = "=A1";*/
        }

        private void OnCellChanged(object sender, EventArgs e)
        {
            SpreadsheetEngine.SpreadsheetCell changedCell = sender as SpreadsheetEngine.SpreadsheetCell;
            if (changedCell == null)
                return;
            dataGridView1.Rows[changedCell.Row].Cells[changedCell.Column].Value = changedCell.CellValue; 
        }

        public static string GetColFromIndex(int colNum)
        { //returns Excel like column names (A, ..., AA, ..., AAA, ...)
            int dividend = colNum;
            string colName = String.Empty;
            int mod;
            while (dividend > 0)
            {
                mod = (dividend - 1) % 26;
                colName = Convert.ToChar(65 + mod).ToString() + colName;
                dividend = (int)((dividend - mod) / 26);
            }
            return colName;
        }

        public static int GetColNumFromName(string colName)
        { //returns column number from Excel style column name
            char[] chars = colName.ToUpperInvariant().ToCharArray();
            int sum = 0;
            for (int i = 0; i < chars.Length; i++)
            {
                sum *= 26;
                sum += (chars[i] - 'A' + 1);
            }
            return sum;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        { //test runner
            Random rand = new Random();
            int row = 0, col = 0; 
            for (int i = 0; i < 50; i++)
            {
                row = rand.Next(0, 49);
                col = rand.Next(2, 25); //keeps first two rows clear
                mainSheet.CellGrid[row, col].CellText = "Hello Friend"; 
            }
            for (int i = 0; i < 50; i++)
            {
                mainSheet.CellGrid[i, 1].CellText = "This is cell B" + (i + 1).ToString();
            }
            for (int i = 0; i < 50; i++)
            {
                mainSheet.CellGrid[i, 0].CellText = "=B" + (i + 1).ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {//load button 
            //select the saveSheet in the main project folder,
            //it's a copy of the "test" button output 
            var openFile = new OpenFileDialog(); 

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                mainSheet.sheet_clear();

                Stream infile = new FileStream(openFile.FileName, FileMode.Open, FileAccess.Read);
                mainSheet.sheet_load(infile);
                infile.Dispose(); 
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {//save button
            var saveFile = new SaveFileDialog(); 

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                Stream outfile = new FileStream(saveFile.FileName, FileMode.Create, FileAccess.Write);
                mainSheet.sheet_save(outfile);
                outfile.Dispose(); 
            }
        }
    }
}
