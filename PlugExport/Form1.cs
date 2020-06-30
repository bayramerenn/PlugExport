using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace PlugExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            

            var folders = Directory.GetFiles(@"C:\transfer").ToList();
            List<Item> items = new List<Item>();

            foreach (var item in folders)
            {
                string text = File.ReadAllText(item);
                string fileName = Path.GetFileNameWithoutExtension(item);
                String[] s = { " ", "\r\n","\0" };

                var texts = text.Split(s, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < texts.Length; i += 2)
                {
                    items.Add(new Item
                    {
                        Barcode = texts[i],
                        Qty = int.Parse(texts[i + 1]),
                        LineDesc = fileName
                    });
                };
            }
            CreateExcel(items);
        }

        public void CreateExcel(List<Item> items)
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Barcode";
            xlWorkSheet.Cells[1, 2] = "ItemTypeCode";
            xlWorkSheet.Cells[1, 3] = "ItemCode";
            xlWorkSheet.Cells[1, 4] = "ColorCode";
            xlWorkSheet.Cells[1, 5] = "ItemDim1Code";
            xlWorkSheet.Cells[1, 6] = "ItemDim2Code";
            xlWorkSheet.Cells[1, 7] = "ItemDim3Code";
            xlWorkSheet.Cells[1, 8] = "Qty1";
            xlWorkSheet.Cells[1, 9] = "Qty2";
            xlWorkSheet.Cells[1, 10] = "LineDescription";
            xlWorkSheet.Cells[1, 11] = "PriceCurrencyCode";
            xlWorkSheet.Cells[1, 12] = "PriceExchangeRate";
            xlWorkSheet.Cells[1, 13] = "Price";
            xlWorkSheet.Cells[1, 14] = "PriceVI";
            xlWorkSheet.Cells[1, 15] = "DocCurrencyCode";
            xlWorkSheet.Cells[1, 16] = "LDisRate1";
            xlWorkSheet.Cells[1, 17] = "LDisRate2";
            xlWorkSheet.Cells[1, 18] = "LDisRate3";
            xlWorkSheet.Cells[1, 19] = "LDisRate4";
            xlWorkSheet.Cells[1, 20] = "LDisRate5";
            xlWorkSheet.Cells[1, 21] = "UnitOfMeasureCode";
            xlWorkSheet.Cells[1, 22] = "PaymentPlanCode";
            xlWorkSheet.Cells[1, 23] = "SalespersonCode";
            xlWorkSheet.Cells[1, 24] = "CostCenterCode";
            xlWorkSheet.Cells[1, 25] = "ITAtt01";
            xlWorkSheet.Cells[1, 26] = "ITAtt02";
            xlWorkSheet.Cells[1, 27] = "ITAtt03";
            xlWorkSheet.Cells[1, 28] = "ITAtt04";
            xlWorkSheet.Cells[1, 29] = "ITAtt05";
            xlWorkSheet.Cells[1, 30] = "LineSumID";
            xlWorkSheet.Cells[1, 31] = "LotCode";
            xlWorkSheet.Cells[1, 32] = "ImportFileNumber";
            xlWorkSheet.Cells[1, 33] = "ExportFileNumber";
            xlWorkSheet.Cells[1, 34] = "ProductSerialNumber";
            xlWorkSheet.Cells[1, 35] = "PurchasePlanCode";
            xlWorkSheet.Cells[1, 36] = "BatchCode";
            xlWorkSheet.Cells[1, 37] = "GLTypeCode";


            for (int i = 1; i < items.Count; i++)
            {
                xlWorkSheet.Cells[i + 1, 1] = items[i].Barcode;
                xlWorkSheet.Cells[i + 1, 8] = items[i].Qty;
                xlWorkSheet.Cells[i + 1, 10] = items[i].LineDesc;
            }

            try
            {
                File.Delete(@"C:\transfer\excel\list.xls");

                xlWorkBook.SaveAs(@"C:\transfer\excel\list.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                MessageBox.Show("Excel dosyasına oluşturuldu.");

            }
            catch (Exception)
            {

                MessageBox.Show("Excel Dosyanız açık");
            }
          

          
           

         
        }
        public class Item
        {
            public string Barcode { get; set; }
            public int Qty { get; set; }
            public string LineDesc { get; set; }
        }
    }
}