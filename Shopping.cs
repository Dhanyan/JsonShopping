using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace jsonCreate
{
    class Shopping
}
class Shopping
{
    public string productType;
    public enum min_price;
    public enum max_price;
    public string City;
    public string Color;
}
static void Main(string[] args)
{
    Shopping shp = new Shopping();
    string JSONresult = Jsonconvert.SerializeObjetc(shp);
    string path = @"C:\json\shopping.json";
    using (var tw = new StreamWriter(path, true))
    {
        tw.WriteLine(JSONresult.Tostring());
        tw.Close();
    }
}