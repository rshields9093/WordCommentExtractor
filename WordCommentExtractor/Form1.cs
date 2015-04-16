using System;
using System.ComponentModel;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordCommentExtractor
{
    public partial class Form1 : Form
    {
        //TODO Refactor such that open file dialog closes sooner
        //TODO Refactor into several discrete methods
        //TODO Study proper program flow for a project like this (i.e. main is hardly used)
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void btnOpenWordFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select Word Document";
            openFileDialog1.Filter = "All files (*.*)|*.*";
            openFileDialog1.Multiselect = false;
            openFileDialog1.ShowDialog();           
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            GetComments(openFileDialog1.FileName);
        }

        private void GetComments(string file)
        {
            object misValue = System.Reflection.Missing.Value;

            //open Word document and get Comments collection
            Word.Application winword = new Word.Application(); //New instance Word
            Word.Document doc = winword.Documents.Open(file); //Open the document and get reference to it
            winword.Visible = false;
            winword.WindowState = Word.WdWindowState.wdWindowStateMinimize;
            Word.Comments comments = doc.Comments; //get Comments collection from document

            //check if there are any comments in the Comments collection
            if (comments.Count == 0) {
                string msg = "There were no comments found in this Word document.  You may open another Word file or simply exit the program.";
                DialogResult result = MessageBox.Show(msg,"Notice!",MessageBoxButtons.OK);
                winword.ActiveDocument.Close();
                winword.Quit();

                return;
            }

            // Start array with the header row
            int numRows = comments.Count + 1;
            int numCols = 5;
            var arrData = new object[numRows, numCols];

            //Set header labels
            arrData[0, 0] = "Date";
            arrData[0, 1] = "Submitted By";
            arrData[0, 2] = "Document";
            arrData[0, 3] = "Page#";
            arrData[0, 4] = "Comments";

            //iterate through all comments and fill in the rest of the array
            int row = 1;
            foreach (Microsoft.Office.Interop.Word.Comment comment in comments)
            {
                string docName = comment.Range.Document.Name;
                string reviewer = comment.Author;
                string date = comment.Date.ToString();
                string commentText = comment.Range.Text;
                Word.Range commentRange = comment.Range.GoTo();
                int pageNumber = commentRange.get_Information(Word.WdInformation.wdActiveEndPageNumber) - 1;

                arrData[row, 0] = date;
                arrData[row, 1] = reviewer;
                arrData[row, 2] = docName;
                arrData[row, 3] = pageNumber;
                arrData[row, 4] = commentText;

                row++;   
            }

            winword.ActiveDocument.Close();
            winword.Quit();
            
            //open new Excel Spreadsheet
            Excel.Application xlApp = new Excel.Application(); //new instance of Excel
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false;

            //check that it succeeded in loading the Excel application
            if (xlApp == null) { MessageBox.Show("Excel is not properly installed!!"); return; }

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(Type.Missing); //New workbook
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //New Worksheet
            xlWorkSheet.Name = "CANES SW2 Doc Reviews";

            //format cells
            ((Excel.Range)xlWorkSheet.Cells[1,1]).EntireColumn.ColumnWidth = 17;
            ((Excel.Range)xlWorkSheet.Cells[1,2]).EntireColumn.ColumnWidth = 25;
            ((Excel.Range)xlWorkSheet.Cells[1,3]).EntireColumn.ColumnWidth = 32;
            ((Excel.Range)xlWorkSheet.Cells[1,4]).EntireColumn.ColumnWidth = 7;
            ((Excel.Range)xlWorkSheet.Cells[1,5]).EntireColumn.ColumnWidth = 105;
            ((Excel.Range)xlWorkSheet.Cells[1,1]).EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1, 1]).EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,2]).EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,2]).EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,3]).EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)xlWorkSheet.Cells[1,3]).EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,4]).EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,4]).EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,5]).EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,1]).EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)xlWorkSheet.Cells[1,1]).EntireRow.Font.Bold = true;

            //set the write range
            var startCell = (Excel.Range)xlWorkSheet.Cells[1,1]; //set beginning of range
            var endCell = (Excel.Range)xlWorkSheet.Cells[numRows,numCols]; //set end of range
            var writeRange = xlWorkSheet.Range[startCell,endCell]; //set the area to write to
            
            //write array to Excel
            writeRange.WrapText = true;
            writeRange.Value2 = arrData;  //write the array to the Excel worksheet 

            //show Excel
            xlApp.Visible = true;
            xlApp.WindowState = Excel.XlWindowState.xlMaximized;
            xlApp.ActiveWindow.Activate();
            xlWorkBook.Saved = false;
                       
            this.Activate();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
