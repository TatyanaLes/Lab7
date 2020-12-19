using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }
//Событие нажатия на область для рисунка
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                pictureBox1.Image = Clipboard.GetImage();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
//Событие нажатия на первую кнопку "Выполнить"
        private void button1_Click(object sender, EventArgs e)
        {
            Word.Application winWord = new Word.Application();
            try
            {
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Multiselect = false,
                    DefaultExt = "*.doc;*.docx",
                    Filter = "Microsoft Word (*.doc*)|*.doc*",
                    Title = "Выберите документ Word"
                };
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string xlFileName = ofd.FileName;

                Word.Document document = winWord.Documents.Open(xlFileName);

                Clipboard.SetImage(pictureBox1.Image);
                winWord.ActiveDocument.Paragraphs[1].Range.Paste();
                winWord.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                winWord.Quit();
            }
        }
//Событие нажатия на вторую кнопку "Выполнить"
        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application winExcel = new Excel.Application();
            try
            {
                OpenFileDialog ofd1 = new OpenFileDialog
                {
                    Multiselect = false,
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "Microsoft Excel (*.xls*)|*.xls*",
                    Title = "Выберите документ Excel №1"
                };
                if (ofd1.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string xlFile1Name = ofd1.FileName;

                OpenFileDialog ofd2 = new OpenFileDialog
                {
                    Multiselect = false,
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "Microsoft Excel (*.xls*)|*.xls*",
                    Title = "Выберите документ Excel №2"
                };
                if (ofd2.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string xlFile2Name = ofd2.FileName;

                Excel.Workbook workbook1;
                Excel.Workbook workbook2;
                Excel.Workbook workbookFinal;
                Excel.Worksheet worksheet1;
                Excel.Worksheet worksheet2;
                Excel.Worksheet worksheetFinal;

                workbook1 = winExcel.Workbooks.Open(xlFile1Name);
                worksheet1 = (Excel.Worksheet)workbook1.Sheets[1];
                workbook2 = winExcel.Workbooks.Open(xlFile2Name);
                worksheet2 = (Excel.Worksheet)workbook2.Sheets[1];
                workbookFinal = winExcel.Workbooks.Add();
                worksheetFinal = (Excel.Worksheet)workbookFinal.Sheets[1];

                int iLastRow = worksheet1.Cells[worksheet1.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
                var arrData = (object[,])worksheet1.Range["A1:Z" + iLastRow].Value;
                int upBound = arrData.GetUpperBound(0) + 1;
                worksheetFinal.Range["A1"].get_Resize(arrData.GetUpperBound(0), arrData.GetUpperBound(1)).Value = arrData;

                iLastRow = worksheet2.Cells[worksheet1.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
                arrData = (object[,])worksheet2.Range["A1:Z" + iLastRow].Value;
                worksheetFinal.Range["A" + upBound].get_Resize(arrData.GetUpperBound(0), arrData.GetUpperBound(1)).Value = arrData;
                SaveFileDialog sfd = new SaveFileDialog()
                {
                    //DefaultExt = "*.xls;*.xlsx",
                    Filter = "Microsoft Excel (*.xls*)|*.xls*",
                    Title = "Выберите место для сохранения"
                };
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    workbookFinal.SaveAs(sfd.FileName);
                }
                winExcel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                winExcel.Quit();
            }
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }
    }
}