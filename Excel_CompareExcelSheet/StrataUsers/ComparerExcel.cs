using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace StrataUsers
{
   class ComparerExcel
    {
        /*
        public static void comparer(string[] args)()



         
            
            {
                List<string> listA = new List<string>() { "a", "b", "c" };
                List<string> listB = new List<string>() { "a", "c", "b" };

                var result = listB.Select((b, index) =>
                    (index == listA.IndexOf(b)) ? b : "");
            }
        }
        */
            //internal bool DoIdsMatchThoseFromXml(List<string> Ids, List<string> XmlIds)
            //{
            //    return
            //        Ids.Count == XmlIds.Count &&
            //        Ids.All(XmlIds.Contains) &&
            //        XmlIds.All(Ids.Contains);


            //}



            //Method to compare two list of string
            //private   List<string> Contains(List<string> CurrentUses, List<string> HREmployeeList)
            //{
            //    List<string> result = new List<string>();

            //    result.AddRange(CurrentUses.Except(HREmployeeList, StringComparer.OrdinalIgnoreCase));
            //    result.AddRange(HREmployeeList.Except(CurrentUses, StringComparer.OrdinalIgnoreCase));

            //    return result;



            //}


            //var workbook1 = new XLWorkbook(@"C:\Automation\Judy_Data\Strata users and HR Employees April 18.xlsx");
            ////var workbook2 = new XLWorkbook(@"workbook2.xlsx");
            //var worksheet1 = workbook.Worksheet("Current User List");
            //var worksheet2 = workbook.Worksheet("HR Employee List");

            //var listSheet1 = new List<IXLRow>(); // list of Rows
            //var listSheet2 = new List<IXLRow>();

            //// puts all UsedRows (including "headers") from sheet1 into a list of rows
            //using (var rows = worksheet1.RowsUsed())
            //{
            //    foreach (var row in rows)
            //    {
            //        listSheet1.Add(row);
            //    }
            //}

            //using (var rows = worksheet2.RowsUsed())
            //{
            //    foreach (var row in rows)
            //    {
            //        listSheet2.Add(row);
            //    }
            //}

            //IEqualityComparer comparer = new XLRowComparer(); // you have to implement your own comparer here. there's a lot of tutorials/samples out there

            //var uniqueIdList = listSheet1.Intersect(listSheet2, comparer).ToList(); // in this case I'd use intersect instead of except which returns the IDs provided in sheet1 and sheet2

        

    }
}
