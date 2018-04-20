using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Globalization;

namespace ObjectToJsonJsonToObject
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Your code goes here
            ExcelSpreadOperation();
            //Demo();
            //Demo2();
            Console.WriteLine("Hello, world!");
        }

        private static List<Dictionary<string, string>> Sort(IEnumerable<Dictionary<string, string>> data, string orderByString)
        {
            var orderBy = orderByString.Split(',').Select(
                 s => s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                  .Select(a =>
                  {
                      var desc = a.Length == 2 && a[1].ToLower() == "desc" ? true : false;
                      return new { Field = a[0], Descending = desc };
                  })
                  .ToList();
            if (orderBy.Count == 0)
                return data.ToList();
            // First one is OrderBy or OrderByDescending.
            IOrderedEnumerable<Dictionary<string, string>> ordered = orderBy[0].Descending ? data.OrderByDescending(d => d[orderBy[0].Field]) : data.OrderBy(d => d[orderBy[0].Field]);
            for (int i = 1; i < orderBy.Count; i++)
            {
                // Rest are ThenBy or ThenByDescending.
                var orderClause = orderBy[i];
                ordered =
                   orderBy[i].Descending ? ordered.ThenByDescending(d => d[orderClause.Field]) : ordered.ThenBy(d => d[orderClause.Field]);
            }
            return ordered.ToList();
        }

        public static void Demo()
        {
            string jsonString = @"[{'COL': 'Aaron','COL1': 'Palmer','COL2': 'Nissim','COL3': 'Cameron'},{'COL': 'Ira','COL1': 'Philip','COL2': 'Henry','COL3': 'Peter'},{'COL': 'Singh','COL1': 'Pratap','COL2': 'Bernard','COL3': 'Lane'}]";

            var list = JsonConvert.DeserializeObject<IEnumerable<Dictionary<string, string>>>(jsonString);

            var colIndex = new Dictionary<int, string>();

            //strData = pRow->GetAt(iCol);
            var pRow = list.ToList()[0];
            int cnt = 0;
            string value;
            if (pRow.TryGetValue("COL", out value))
            {

            }
            foreach (KeyValuePair<string, string> kvp in pRow)
            {
                colIndex.Add(cnt++, kvp.Key);
            }

            string orderbyS = "COL1,COL asc,COL2 DESC";
            //IEnumerable<Dictionary<string, string>> listy = Sort(list, orderbyS);
        }

        public static void ExcelSpreadOperation()
        {
            ExcelOperation excelOperation = new ExcelOperation();
            string path = @"C:\Users\Pratap\Desktop\EIGS_Sample3.xlsx";

            excelOperation.OpenExcelBook(path);

            excelOperation.SetWorksheet("Sheet1");

            //excelOperation.DoSomeOperation();

            //excelOperation.DoSomeOperationFour();

            //excelOperation.DoSomeOperationTwo();


            //excelOperation.SetWorksheet("Sheet2");

            // excelOperation.DoSomeOperationThree();

            //excelOperation.ProtectUnprotectSheet(null);

            // excelOperation.MergeAreaCheck(null);

            // excelOperation.GetAddrCheck();

            DateTimeCheck();




        }

        public static void Demo2()
        {
            Decimal number = 6.06M;
            Console.WriteLine(number.ToString() == "6.06");

            string test = "9,000,000,000";

            test.Replace(",", "");
            int index = -1;
            while ((index = test.IndexOf(',')) > -1)
            {
                test = test.Remove(index, 1);
            }
        }

        public static void DateTimeCheck()
        {

            ParseDateTime("Friday, April 10, 2009", 0, new CultureInfo("ja-JP"), out DateTime date);
            ParseDateTime("Friday, April 10, 2009", 1, new CultureInfo("ja-JP"), out DateTime dateOne);
            ParseDateTime("Jan 1, 2009", 1, new CultureInfo("ja-JP"), out DateTime dateTwo);


        }

        public static bool ParseDateTime(string strDate, int dwFlags, CultureInfo cultureinfo, out DateTime date)
        {
            date = new DateTime();

            try
            {
                if (dwFlags == 0)
                {
                    string format = cultureinfo.DateTimeFormat.ShortDatePattern;
                    date = DateTime.Parse(strDate, cultureinfo);
                }
                else
                {
                    string format = cultureinfo.DateTimeFormat.ShortTimePattern;
                    date = DateTime.Parse(strDate, cultureinfo);
                }
            }
            catch (ArgumentNullException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            return true;

        }

    }
}



