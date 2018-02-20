using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Linq; 

namespace SpreadsheetEngine
{
    public class Cell : INotifyPropertyChanged
    {
        private int RowIndex;
        private int ColumnIndex;
        private string Text;
        protected string Value;
        private string Name; //saves column/row name

        public event PropertyChangedEventHandler PropertyChanged;

        public Cell(int row, int col)
        {
            RowIndex = row;
            ColumnIndex = col;

            Name += Convert.ToChar('A' + col);
            Name += (row + 1).ToString(); 
        }

        public int Row
        {
            get { return RowIndex; }
        }
        public int Column
        {
            get { return ColumnIndex; }
        }

        public string CellText
        {
            get { return Text; }
            set
            {
                if (value != Text)
                {
                    Text = value;
                    OnPropertyChanged("CellText");
                }
            }
        }

        public string CellValue
        {
            get { return Value; }
        }

        public string CellName
        {
            get { return Name;  }
        }

        public void cell_clear() //reset cell value 
        {
            CellText = "";  
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }

    public class SpreadsheetCell : Cell
    {
        public SpreadsheetCell(int row, int col) : base(row, col) { }

        public string SetValue
        {
            set
            {
                Value = value;
            }
        }
    }

    public class Spreadsheet
    {
        private int Rows;
        private int Cols;
        public SpreadsheetCell[,] CellGrid;
        public event PropertyChangedEventHandler CellChanged;

        public Spreadsheet(int rows, int cols)
        {
            Rows = rows;
            Cols = cols;

            SpreadsheetCell[,] newGrid = new SpreadsheetCell[Rows, Cols];
            CellGrid = newGrid;

            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    CellGrid[i, j] = new SpreadsheetCell(i, j);
                    CellGrid[i, j].PropertyChanged += CellPropertyChanged;
                }
            }
        }

        public int RowCount
        {
            get { return Rows; }
            set { Rows = value; }
        }
        public int ColumnCount
        {
            get { return Cols; }
            set { Cols = value; }
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
                dividend = ((dividend - mod) / 26);
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

        public SpreadsheetCell GetCell(int row, int column)
        {
            return CellGrid[row, column];
        }
        public SpreadsheetCell GetCell(string address)
        {
            char colIndex = address[0];
            int rowIndex;
            SpreadsheetCell result;

            if (!Char.IsLetter(colIndex))
                return null;

            if (!int.TryParse(address.Substring(1), out rowIndex))
                return null; //if can't find row index

            try
            {
                result = GetCell(rowIndex - 1, colIndex - 'A'); 
            }
            catch (Exception c)
            {
                return null; 
            }

            return result; 
        }

        protected void CellPropertyChanged(object sender, EventArgs e)
        {
            var targetCell = sender as SpreadsheetCell;
            if (targetCell == null)
                return;
            string text = targetCell.CellText;
            if (text == "")
            {
                targetCell.SetValue = text; 
            }
            else if (text[0] == '=')
            {
                string columnStr = text[1].ToString();
                int columnNum = GetColNumFromName(columnStr) - 1;
                int rowNum = Convert.ToInt32(text.Substring(2)) - 1;
                if (GetCell(rowNum, columnNum).CellText != null)
                {
                    targetCell.SetValue = GetCell(rowNum, columnNum).CellText;
                    CellChanged(targetCell, new PropertyChangedEventArgs("Value"));
                }
                else
                {
                    targetCell.SetValue = targetCell.CellText;
                    CellChanged(targetCell, new PropertyChangedEventArgs("Value"));
                }  
            }
            else
            {
                targetCell.SetValue = targetCell.CellText;
                CellChanged(targetCell, new PropertyChangedEventArgs("Value")); 
            }
        }

        public void sheet_save(Stream outfile)
        {
            XmlWriter author = XmlWriter.Create(outfile);
            //begin XML spreadsheet file
            author.WriteStartElement("spreadsheet"); 

            foreach (Cell cell in CellGrid)
            {
                if (cell.CellText != "" || cell.CellValue != "") //if cell isn't default blank, save
                {
                    author.WriteStartElement("cell");
                    author.WriteAttributeString("name", cell.CellName);
                    author.WriteElementString("text", cell.CellText);
                    //author.WriteElementString("bgcolor", cell.CellColor) 
                    //I don't know how to add support for the cell color on the
                    //spreadsheet, but I understand how the XML format would work here,
                    //we were just never asked to support color throughout the project
                    author.WriteEndElement(); 
                }
            }
            author.WriteEndElement();
            author.Close(); 
        }
        public void sheet_load(Stream infile)
        {
            XDocument newfile = XDocument.Load(infile);

            foreach (XElement tag in newfile.Root.Elements("cell"))
            {
                Cell newcell = GetCell(tag.Attribute("name").Value.ToString());

                if (tag.Element("text") != null)
                {
                    newcell.CellText = tag.Element("text").Value.ToString();
                    /*
                    will not dynamically update previously written cells 
                    Ex: if A1 is written as "=B1" before B1 is written, 
                    it will appear blank on the final spreadsheet since 
                    B1 was initialized with its Text == null
                 */
                }
                //if (tag.Element("bgcolor") != null)
                    //see explanation above
            }
        }

        public void sheet_clear()
        {
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    if (CellGrid[i, j].CellText != "" || CellGrid[i, j].CellValue != "")
                    {
                        CellGrid[i, j].cell_clear(); 
                    }
                }
            }
        }
    }

    //Expression Tree Classes
    public class ExpTree
    {
        private Dictionary<string, double> vDict = new Dictionary<string, double>();
        private Node root;
        private string expressionString;
        public static char[] opcodes = { '+', '-', '*', '/' }; 

        private abstract class Node
        {
            protected string Name;
            protected double Value;

            public Node Left;
            public Node Right; 

            public string get_name()
            {
                return Name; 
            }
            public double get_val()
            {
                return Value; 
            }
            public void set_val(double num)
            {
                Value = num; 
            }
        }

        private class OpNode : Node
        {
            public OpNode(string name, Node left, Node right)
            { //constructor for Operand Nodes
                Name = name;
                Left = left;
                Right = right; 
            }
        }
        private class VarNode: Node
        {
            public VarNode(string name)
            { //Variable Node constructor (only uses name in construction) 
                Name = name; 
            }
        }
        private class ConstNode : Node
        {
            public ConstNode(double number)
            { //Constant Node constructor (only uses Value) 
                Value = number; 
            }
        }

        public string ExpressionString
        {
            get { return expressionString; }
            set
            {
                expressionString = value;
                vDict.Clear();
                root = Compile(expressionString); 
            }
        }

        private int FindOp(string exp)
        { //returns index for operator symbol from user expression
            for (int i = 0; i < exp.Length; i++)
            {
                if ((exp[i] == '+') || (exp[i] == '-') || (exp[i] == '*') || (exp[i] == '/'))
                {
                    return i; //return index of expression location 
                }
            }
            return -1; //no expression found 
        }

        private Node Compile(string newExpression)
        {
            int parensCount = 0; 

            if ((newExpression == "") || (expressionString == null))
            {
                return null; 
            }

            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            //HW6 PARENTHESES SUPPORT 
            if (newExpression[0] == '(') //if expression starts with parentheses
            {
                for (int i = 0; i < newExpression.Length; i++)
                {
                    if (newExpression[i] == '(')
                    {
                        parensCount++; 
                    }
                    else if (newExpression[i] == ')')
                    {
                        parensCount--; 

                        if (parensCount == 0) //if parentheses are balanced
                        {
                            if (newExpression.Length - 1 != i)
                                break; 
                            else //if at end of expression bounded by parentheses 
                            { //cut parentheses, evaulate substrings 
                                return Compile(newExpression.Substring(1, newExpression.Length - 2)); 
                            }
                        }
                    }
                }
            }
            //HW6 PARENTHESES SUPPORT 
            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            foreach (char op in opcodes)
            {//opcodes follows order of operations
                //creates opnodes with */ closer to leaves, +- closer to root
                Node n = OpCompile(newExpression, op);
                if (n != null)
                    return n; 
            }

            int opIndex = FindOp(newExpression);
            int constInt; 
            if (opIndex == -1)
            {
                if (int.TryParse(newExpression, out constInt))
                {
                    return new ConstNode(constInt); 
                }
                else
                {
                    vDict.Add(newExpression, 0);
                    return new VarNode(newExpression); 
                }
            }

            Node LeftNode = Compile(newExpression.Substring(0, opIndex));
            Node RightNode = Compile(newExpression.Substring(opIndex + 1));

            return new OpNode(newExpression[opIndex].ToString(), LeftNode, RightNode); 
        }

        private Node OpCompile(string expression, char op)
        {
            int i = expression.Length - 1;
            int parensCount = 0; 
            bool quitter = false;

            while (!quitter)
            {
                if (expression[i] == '(')
                    parensCount++;
                else if (expression[i] == ')')
                    parensCount--; 

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                //HW6 PARENTHESES SUPPORT
                if (parensCount == 0 && expression[i] == op)
                { //if not currently inside parenthesis 
                    return new OpNode(op.ToString(), Compile(expression.Substring(0, i)), Compile(expression.Substring(i + 1))); 
                }
                //HW6 PARENTHESES SUPPORT
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                else
                {
                    if (i == 0)
                        quitter = true;
                    i--; 
                }
            }

            return null; 
        }

        //Calculates an Expression Tree given it's root node 
        private double Evaluate(Node n)
        {
            if (n == null)
                return -1;

            ConstNode constTest = n as ConstNode;
            if (constTest != null)
                return constTest.get_val();

            VarNode varTest = n as VarNode; 
            if (varTest != null)
            {
                try
                {
                    return vDict[n.get_name()]; 
                }
                catch (Exception c)
                {
                    Console.WriteLine("Key Not Found, Undefined Behavior"); 
                }
            }

            OpNode opTest = n as OpNode;
            if (opTest != null)
            {
                if (opTest.get_name() == "+")
                    return (Evaluate(opTest.Left) + Evaluate(opTest.Right));

                else if (opTest.get_name() == "-")
                    return (Evaluate(opTest.Left) - Evaluate(opTest.Right));

                else if (opTest.get_name() == "*")
                    return (Evaluate(opTest.Left) * Evaluate(opTest.Right));

                else if (opTest.get_name() == "/")
                    return (Evaluate(opTest.Left) / Evaluate(opTest.Right));
            }

            return 0.0; //default return value 
        }
        
        public double Eval() //public interface for evaluate code
        {
            return Evaluate(root); 
        }

        public double SetVar(string vName, double vValue)
        {
            try
            {
                vDict[vName] = vValue;
                return vDict[vName]; 
            }
            catch (Exception c)
            {
                Console.WriteLine("Error: Variable not in dictionary");
                return -1; 
            }
        }
    }
}
