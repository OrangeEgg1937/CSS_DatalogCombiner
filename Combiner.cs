using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CsvHelper.Configuration.Attributes;

namespace CSS_DatalogCombiner
{
    class Combiner
    {
        public class Entity
        {
            [Name("Sample:float")] public float Sample { get; set; }
            [Name("Data:float")] public float Data { get; set; }
        }

        Microsoft.Office.Interop.Excel.Application xlApp;
        String PATH = "C:\\Users\\aenkwho\\Desktop\\datalog\\excel_marco\\result.xlsx";
        List<String> importFile = new List<String>();
        String xTitle;
        String yTitle;

        public void setTitle(String x, String y)
        {
            xTitle = x;
            yTitle = y;
        }

        public void ImportFile(String FILE_PATH)
        {
            importFile.Add(FILE_PATH);
        }

        public void DisplayPathNameHandler()
        {
            foreach (string myStringList in importFile)
            {
                Console.WriteLine(myStringList);
            }
        }

        public void CreateExcel()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Time";
            xlWorkSheet.Cells[1, 2] = "Name";

            int i = 2;
            int j = 0;

            foreach (string myStringList in importFile)
            {
                var reader = new StreamReader(myStringList);
                var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
                var records = csv.GetRecords<Entity>();

                foreach (var record in records)
                {
                    xlWorkSheet.Cells[i, 1] = (record.Sample+j*0.5);
                    xlWorkSheet.Cells[i, 2] = record.Data;
                    i++;
                }
                j++;
            }
         
            // Here saving the file in xlsx
            xlWorkBook.SaveAs(PATH, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            // Release .exe for task manager 
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file in "+ PATH);
        }
    }
}
