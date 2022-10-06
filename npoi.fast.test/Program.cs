using npoi.fast;
using System.Data;

namespace npoi.fast.test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "D:\\downloads\\报损商品导入模板.xls";
            List<string[]> list =ExcelUtils.ToList(path,5);
            Console.WriteLine("完成");
        }
    }
}