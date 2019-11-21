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
    class CDLClassicUsersList
    {

        public static List<CDLClassicUsers> ReadExcelForCDLUsers(string tab, string filelocation, int col1, int col2, int col3, int col4)

        {
            List<CDLClassicUsers> AllCDLUsers = new List<CDLClassicUsers>();


            var newFile = new ExcelPackage(new FileInfo(@filelocation));


            ExcelWorksheet CDLClassicUsersList = newFile.Workbook.Worksheets[tab];



            for (int i = CDLClassicUsersList.Dimension.Start.Row;
         i <= CDLClassicUsersList.Dimension.End.Row;
         i++)
            {
                CDLClassicUsers myCDLClassicUser = new CDLClassicUsers();

                if (CDLClassicUsersList.Cells[i, col1].Value == null)
                {
                    myCDLClassicUser.Op_Code = "";
                }
                else
                {
                    myCDLClassicUser.Op_Code = CDLClassicUsersList.Cells[i, col1].Value.ToString();
                }

                if (CDLClassicUsersList.Cells[i, col2].Value == null)
                {
                    myCDLClassicUser.Name = "";
                }
                else
                {
                    myCDLClassicUser.Name = CDLClassicUsersList.Cells[i, col2].Value.ToString();
                }

                if (CDLClassicUsersList.Cells[i, col3].Value == null)
                {
                    myCDLClassicUser.Printer = "";
                }
                else
                {
                    myCDLClassicUser.Printer = CDLClassicUsersList.Cells[i, col3].Value.ToString();
                }
                if (CDLClassicUsersList.Cells[i, col4].Value == null)
                {
                    myCDLClassicUser.Department = "";
                }
                else
                {
                    myCDLClassicUser.Department = CDLClassicUsersList.Cells[i, col4].Value.ToString();
                }



              
                AllCDLUsers.Add(myCDLClassicUser);


            }


            //AllRisks.RemoveAt(0);
            //risk.Address_1
            return AllCDLUsers;


        }

        public static void WriteExceptions(List<CDLClassicUsers> exceptionList, string filelocation)
        {
            var newFile = new ExcelPackage(new FileInfo(@filelocation));
            newFile.Workbook.Worksheets.Add("Exceptions");

            ExcelWorksheet exceptions = newFile.Workbook.Worksheets["Exceptions"];

           

                exceptions.Cells[1, 1].Value = "OP CODE";
                exceptions.Cells[1, 2].Value = "NAMES";
                exceptions.Cells[1, 3].Value = "PRINTER";
                exceptions.Cells[1, 4].Value = "DEPARTMENT";


                exceptions.Column(1).Width = 15;
                exceptions.Column(2).Width = 30;
                exceptions.Column(3).Width = 35;
                exceptions.Column(4).Width = 30;

                exceptions.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                exceptions.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Purple);


                  
                exceptions.Row(1).Style.Font.Color.SetColor(Color.White);

                exceptions.Column(1).Style.Font.Size = 12;
                exceptions.Row(1).Style.Font.Size = 14;
                string path = filelocation;
                FileInfo file = new FileInfo(path);

               
                using (var excelFile = new ExcelPackage(file))
                {
                   
                    exceptions.Cells["A1"].LoadFromCollection(Collection: exceptionList, PrintHeaders: true);
                    

                    excelFile.Save();
                }

               

                int row1 = 1;
                foreach (var item in exceptionList)
                {
                    exceptions.Cells[row1, 1].Value = item.Op_Code;
                    exceptions.Cells[row1, 2].Value = item.Name;
                    exceptions.Cells[row1, 3].Value = item.Printer;
                    exceptions.Cells[row1, 4].Value = item.Department;

                   


                    row1++;





                }



                newFile.Save();

            }


        }
    }



