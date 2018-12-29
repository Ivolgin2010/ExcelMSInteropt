using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace ExcelMSInteropt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // параметра передать Missing.Value
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;

            //создаем и инициализируем объекты Excel
            Excel.Application App;
            Excel.Workbook xlsWB;
            Excel.Worksheet xlsSheet;

            App = new Microsoft.Office.Interop.Excel.Application();
            //добавляем в файл Excel книгу. Параметр в данной функции - используемый для создания книги шаблон.
            //если нас устраивает вид по умолчанию, то можно спокойно передавать пустой параметр.
            xlsWB = App.Workbooks.Add(missingValue);
            //и использует из нее
            xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(1);

            List<string[]> rows = new List<string[]>
            {
                new string[] { "1", "2", "3", "4" },
                new string[] { "5", "6", "7", "8" },
                new string[] { "9", "10", "11", "12" }
            };
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 0; j < rows[i].Length; j++)
                {
                    xlsSheet.Cells[i + 1, j + 1] = rows[i][j];
                }
            }

            //у SaveAs масса  параметров, некоторые из которых могут оказаться для вас полезными. 
            //Сейчас мы используем только тип формата Excel12 (Excel 97) и права доступа.
            //Полное описание что и как смотрите на MSDN: 
            // http://msdn.microsoft.com/ru-ru/library/microsoft.office.tools.excel.workbook.saveas.aspx
            xlsWB.SaveAs(@"C:\Users\i.geraskin\source\repos\excel_text.xls",
                         Excel.XlFileFormat.xlExcel12,
                         missingValue,
                         missingValue,
                         missingValue,
                         missingValue,
                         Excel.XlSaveAsAccessMode.xlNoChange,
                         missingValue,
                         missingValue,
                         missingValue,
                         missingValue,
                         missingValue);
            //закрываем книгу                                                                        
            xlsWB.Close(true, missingValue, missingValue);
            //закрываем приложение
            App.Quit();

            MessageBox.Show("Все сохранено", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //уменьшаем счетчики ссылок на COM объекты, что, по идее должно их освободить.
            //почему это не произойдет - читайте ниже ;)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWB);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(App);

        }
    }
}
