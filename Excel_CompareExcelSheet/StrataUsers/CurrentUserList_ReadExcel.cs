using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Drawing;


namespace StrataUsers
{
    public class CurrentUserList_ReadExcel
    {

        public static List<CurrentUsers> ReadExcelForCurrentUsers(string tab, string filelocation, int col)
        {
            List<CurrentUsers> AllCurrentUsers = new List<CurrentUsers>();


            var newFile = new ExcelPackage(new FileInfo(@filelocation));


            ExcelWorksheet CurrentUsersList = newFile.Workbook.Worksheets[tab];



            for (int i = CurrentUsersList.Dimension.Start.Row;
         i <= CurrentUsersList.Dimension.End.Row;
         i++)
            {
                CurrentUsers myOGIUser = new CurrentUsers();

                myOGIUser.Name = CurrentUsersList.Cells[i, col].Value.ToString();
                AllCurrentUsers.Add(myOGIUser);

            }
           
            //AllRisks.RemoveAt(0);
            //risk.Address_1
            return AllCurrentUsers;

        }






        public static void WriteExceptions(List<string> exceptionList, string filelocation)
        {
            var newFile = new ExcelPackage(new FileInfo(@filelocation));
            newFile.Workbook.Worksheets.Add("Exceptions");

            ExcelWorksheet exceptions = newFile.Workbook.Worksheets["Exceptions"];

           


            

            int total = exceptionList.Count;

            for (int row = 2; row < total; row++)
            {

                exceptions.Cells[1, 1].Value = "EXCEPTIONS NAMES";

               

                exceptions.Column(1).Width = 35;



                exceptions.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                exceptions.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Purple);

               
                exceptions.Row(1).Style.Font.Color.SetColor(Color.White);
                
                exceptions.Column(1).Style.Font.Size = 12;
                exceptions.Row(1).Style.Font.Size = 14;


                exceptions.Cells[row, 1].Value = exceptionList[row];



               
            }



            newFile.Save();

        }
        

        }

    }


    

