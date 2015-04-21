using System;
using System.ComponentModel;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

//TODO Refactor into several discrete methods
//TODO Add a progress bar
//TODO Open Word docs in readonly (to prevent warning if the file is already open and locked)

namespace WordCommentExtractor
{
    public partial class Form1 : Form
    {        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) {
            txtStatus.Text = "Waiting for user to select some Word files...";
        }

        private void btnOpenWordFile_Click(object sender, EventArgs e)
        {
            //Open file dialog and store the returned value
            DialogResult result = openFileDialog1.ShowDialog();
            
            //If Open Button was pressed
            if (result == DialogResult.OK)
            {
                foreach (var item in openFileDialog1.FileNames)
                {
                    GetComments(item);
                }
            }

            txtStatus.AppendText("\r\nFinished processing all Word Files.");
            txtStatus.AppendText("\r\nYou may exit now or select more Word Files.");
            this.Activate();                       
        }

        private void GetComments(string file)
        {

            txtStatus.AppendText("\r\nOpening Word file: " + file);
            //open Word document and get Comments collection
            Word.Application winword = new Word.Application(); //New instance of Word
            Word.Document doc = winword.Documents.Open(file); //Open the document and get reference to it
            winword.Visible = false;
            winword.WindowState = Word.WdWindowState.wdWindowStateMinimize;
            Word.Comments comments = doc.Comments; //get Comments collection from document

            //check if there are any comments in the Comments collection
            if (comments.Count == 0) {
                txtStatus.AppendText("\r\nNOTE: There were no comments found in " + file + ". Moving on to next file (if any).");
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
            txtStatus.AppendText("\r\nProcessing comments in " + file);
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
            txtStatus.AppendText("\r\nFinished processing all comments in " + file);

            winword.ActiveDocument.Close();
            winword.Quit();
            
            //open new Excel Spreadsheet
            txtStatus.AppendText("\r\nCreating new Excel file...");
            Excel.Application xlApp = new Excel.Application(); //new instance of Excel
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false;

            //check that it succeeded in loading the Excel application
            if (xlApp == null) { MessageBox.Show("Excel is not properly installed!!"); return; }

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(Type.Missing); //New workbook
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //New Worksheet
            xlWorkSheet.Name = "CANES SW2 Doc Reviews";

            txtStatus.AppendText("Formatting Excel spreadsheet...");
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

            txtStatus.AppendText("\r\nWriting comments to spreadsheet");
            //write array to Excel
            writeRange.WrapText = true;
            writeRange.Value2 = arrData;  //write the array to the Excel worksheet 

            //show Excel
            xlApp.Visible = true;
            xlApp.WindowState = Excel.XlWindowState.xlMaximized;
            xlApp.ActiveWindow.Activate();
            xlWorkBook.Saved = false;
            
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
