using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace jsonCreate
{
    class Shopping
    {
        public string productType;
        public string properties;
        public int price;
        public string storeaddress;

        static void Main(string[] args)
        {
            Shopping shp = new Shopping();
            shoppingData shpData = new shoppingData();
            shp = shpData.ShoppingDataOps();
            string JSONresult = Newtonsoft.Json.JsonConvert.SerializeObject(shp);
            string path = @"C:\json\shopping.json";
            if (File.Exists(path))
            {
                File.Delete(path);

                using (var tw = new StreamWriter(path, true))
                {
                    tw.WriteLine(JSONresult.ToString());
                    tw.Close();
                }
            }
            else if (!File.Exists(path))
            {
                using (var tw = new StreamWriter(path, true))
                {
                    tw.WriteLine(JSONresult.ToString());
                    tw.Close();
                }
            }
        }

        public class shoppingData
        {

            public Shopping ShoppingDataOps()
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(@"C:\json\d.xlsx", 0);
                xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets["Sheet1"];

                var type = xlWorksheet.Columns.Find("Product type").Cells.Column;
                string producttype = xlWorksheet.Cells[2, type].ToString();

                var properties = xlWorksheet.Columns.Find("Product properties").Cells.Column;
                string productproperty = xlWorksheet.Cells[2, type].ToString();

                var price = xlWorksheet.Columns.Find("Price").Cells.Column;
                int productPrice = Convert.ToInt32(xlWorksheet.Cells[2, type]);

                var address = xlWorksheet.Columns.Find("Store address").Cells.Column;
                string storeAddress = xlWorksheet.Cells[2, type].ToString();

                Shopping shpObj = new Shopping();
                shpObj.productType = producttype;
                shpObj.properties = productproperty;
                shpObj.price = productPrice;
                shpObj.storeaddress = storeAddress;
                return shpObj;
            }
        }
    }

}