using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;


namespace StrataUsers
{
    class HRList_ReadExel
    {
        public static List<HREmployee> ReadExcelForHREmployeeList(string Tab,string filelocation,int col,int col1)
        {
            List<HREmployee> AllHREmployeesList = new List<HREmployee>();
        

            var newFile = new ExcelPackage(new FileInfo(filelocation));



            ExcelWorksheet HREmployeeList = newFile.Workbook.Worksheets[Tab];



            for (int i = HREmployeeList.Dimension.Start.Row;
         i <= HREmployeeList.Dimension.End.Row;
         i++)
            {
                HREmployee myHREmployee = new HREmployee();

                myHREmployee.FirstName = HREmployeeList.Cells[i, col].Value.ToString();
                  myHREmployee.Surname = HREmployeeList.Cells[i, col1].Value.ToString();

                AllHREmployeesList.Add(myHREmployee);

                //for (int j = HREmployeeList.Dimension.Start.Column;
                //         j <= HREmployeeList.Dimension.End.Column;
                //         j++)
                //{
                //    //myHREmployee.FirstName = HREmployeeList.Cells[j, 2].Value.ToString();
                //    //myHREmployee.Surname = HREmployeeList.Cells[j, 3].Value.ToString();
                   
                //}
               
            }
            AllHREmployeesList.RemoveAt(0);

            //AllRisks.RemoveAt(0);
            //risk.Address_1
            return AllHREmployeesList;

        }

        public static List<HREmployee> ReadExcelForHREmployeeList1(string Tab1, string filelocation1,int colm1,int colm2)
        {
            List<HREmployee> AllHREmployeesList = new List<HREmployee>();


            var newFile = new ExcelPackage(new FileInfo(filelocation1));



            ExcelWorksheet HREmployeeList = newFile.Workbook.Worksheets[Tab1];



            for (int i = HREmployeeList.Dimension.Start.Row;
         i <= HREmployeeList.Dimension.End.Row;
         i++)
            {
                HREmployee myHREmployee = new HREmployee();

                myHREmployee.FirstName = HREmployeeList.Cells[i, colm1].Value.ToString();
                myHREmployee.Surname = HREmployeeList.Cells[i, colm2].Value.ToString();

                AllHREmployeesList.Add(myHREmployee);

                //for (int j = HREmployeeList.Dimension.Start.Column;
                //         j <= HREmployeeList.Dimension.End.Column;
                //         j++)
                //{
                //    //myHREmployee.FirstName = HREmployeeList.Cells[j, 2].Value.ToString();
                //    //myHREmployee.Surname = HREmployeeList.Cells[j, 3].Value.ToString();

                //}

            }
            AllHREmployeesList.RemoveAt(0);

            //AllRisks.RemoveAt(0);
            //risk.Address_1
            return AllHREmployeesList;

        }

    }

    }

