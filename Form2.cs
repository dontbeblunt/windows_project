using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;



namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public string dir;
        public string dir1 = "C:\\Users\\Pavel\\Desktop\\WARNING";
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Проверяем наличие файла
           // if (File.Exists(dir+"\\report-4.xls"))
                if (File.Exists(dir))
                {
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                       
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(dir, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                //Очищаем поля
               // textBox1.Clear();

                {
                    //Ищем данные в столбце "С"
                    Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("AA:AA").Find(textBox1.Text);
                   
                    //Добавляем текст из нужных ячеек.
                    textBox2.Text = ObjWorkSheet.get_Range("DP" + range.Row.ToString()).Value2;
                    textBox3.Text = ObjWorkSheet.get_Range("DQ" + range.Row.ToString()).Value2;
                    textBox4.Text = ObjWorkSheet.get_Range("DR" + range.Row.ToString()).Value2;
                    textBox5.Text = ObjWorkSheet.get_Range("BO" + range.Row.ToString()).Value2;
                    textBox6.Text = ObjWorkSheet.get_Range("CU" + range.Row.ToString()).Value2;
                    textBox7.Text = ObjWorkSheet.get_Range("DU" + range.Row.ToString()).Value2;
                    textBox8.Text = ObjWorkSheet.get_Range("DO" + range.Row.ToString()).Value2;
                    textBox10.Text = ObjWorkSheet.get_Range("AI" + range.Row.ToString()).Value2;
                    //это чтобы форма прорисовывалась (не подвисала)...
                    System.Windows.Forms.Application.DoEvents();
                }

                //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                ObjWorkBook.Close();
                ObjExcel.Quit();
                /*int i = 1;
                int k = 0;
                string stolb;
                 excelapp = new Excel.Application();
                excelapp.Visible = false;
                excelappworkbooks = excelapp.Workbooks;
                excelappworkbook = excelapp.Workbooks.Open(Application.StartupPath + @"\template\2.xls",
                          Type.Missing, true, Type.Missing,
        "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
         Type.Missing, Type.Missing);
                excelsheets = excelappworkbook.Worksheets;
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                excelcells = excelworksheet.Cells.Find("q", Missing.Value, Missing.Value, Excel.XlLookAt.xlPart, Missing.Value,
                   Excel.XlSearchDirection.xlNext,
                   Missing.Value, Missing.Value, Missing.Value);

                stolb = Convert.ToString(excelcells.Column);
                strok = Convert.ToString(excelcells.Rows.Row);
                System.Windows.Forms.MessageBox.Show(getAdres(Convert.ToInt32(poisk), (Convert.ToInt32(stolb)) - 1));
                excelapp.Quit();*/
            }
            else
            {
                MessageBox.Show(
                    "Файл не выбран или отсутсвует в каталоге", "Ошибка"
                );
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            /* FolderBrowserDialog DirDialog = new FolderBrowserDialog();
             DirDialog.Description = "Выбор директории";
             DirDialog.SelectedPath = @"C:\";


             if (DirDialog.ShowDialog() == DialogResult.OK)
             {
                 dir = DirDialog.SelectedPath;
             }*/
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                dir = openFileDialog1.FileName;              
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
                Microsoft.Office.Interop.Word.Application ObjWord = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document Doc = ObjWord.Documents.Add(dir + "\\Prilozhenie4.docx");
                Doc.Bookmarks["Name"].Range.Text = textBox2.Text + " " + textBox3.Text + " " + textBox4.Text;
                Doc.Bookmarks["Date_of_Birth"].Range.Text = ", " + textBox5.Text + " г.";
                Doc.Bookmarks["Seriya"].Range.Text = textBox6.Text;
                Doc.Bookmarks["Number"].Range.Text = textBox7.Text;
                Doc.Bookmarks["Type"].Range.Text = textBox8.Text;
                Doc.Bookmarks["Learning_year"].Range.Text = textBox9.Text;
                Doc.Bookmarks["City"].Range.Text = textBox10.Text;
                Doc.Bookmarks["Srok"].Range.Text = textBox11.Text;
                Doc.Bookmarks["Begin_Learning"].Range.Text = textBox12.Text;
                Doc.Bookmarks["End_Learning"].Range.Text = textBox13.Text;
                Doc.Bookmarks["Rector"].Range.Text = comboBox1.Text;
                Doc.Bookmarks["Name2"].Range.Text = comboBox2.Text;
                Doc.Bookmarks["Director"].Range.Text = comboBox3.Text;
                Doc.Bookmarks["Name3"].Range.Text = comboBox4.Text;

                Doc.SaveAs(FileName: dir1 + "\\" + textBox2.Text + "_" + textBox3.Text + "_For_print.docx");
                Doc.Close();
                ObjWord.Quit();
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }
    }
}
