using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using System.Collections.ObjectModel;

namespace StrataUsers
{
    class I90UserAccountList_ReadExcel
    {


            public static List<i90UserAccount> ReadExcelFori90Users(string tab, string filelocation, int col1, int col2, int col3, int col4, int col5)

            {
                List<i90UserAccount> Alli90UserAccount = new List<i90UserAccount>();

                var newFile = new ExcelPackage(new FileInfo(@filelocation));

                ExcelWorksheet i90UserAccountList = newFile.Workbook.Worksheets[tab];

                for (int i = i90UserAccountList.Dimension.Start.Row;
             i <= i90UserAccountList.Dimension.End.Row;
             i++)
                {
                    i90UserAccount myi90UserAccounts = new i90UserAccount();

                   

                    if (i90UserAccountList.Cells[i, col1].Value == null)
                    {
                        myi90UserAccounts.User = "";
                    }
                    else
                    {
                    myi90UserAccounts.User = i90UserAccountList.Cells[i, col1].Value.ToString();
                    }

                    if (i90UserAccountList.Cells[i, col2].Value.ToString() == null)
                    {
                    myi90UserAccounts.Text = "";
                    }
                    else
                    {
                    myi90UserAccounts.Text = i90UserAccountList.Cells[i, col2].Value.ToString();
                    }
                    if (i90UserAccountList.Cells[i, col3].Value == null)
                    {
                    myi90UserAccounts.Status = "";
                    }
                    else
                    {
                    myi90UserAccounts.Status = i90UserAccountList.Cells[i, col3].Value.ToString();
                    }

                    if (i90UserAccountList.Cells[i, col4].Value == null)
                    {
                    myi90UserAccounts.Date_Creation = "";
                    }
                    else
                    {
                    myi90UserAccounts.Date_Creation = i90UserAccountList.Cells[i, col4].Value.ToString();
                    }

                    if (i90UserAccountList.Cells[i, col5].Value.ToString() == null)
                    {
                    myi90UserAccounts.Date_Previous_Sign_on = "";
                    }
                    else
                    {
                    myi90UserAccounts.Date_Previous_Sign_on = i90UserAccountList.Cells[i, col5].Value.ToString();
                    }
                    



                    Alli90UserAccount.Add(myi90UserAccounts);


                }

                //AllRisks.RemoveAt(0);
                //risk.Address_1
                return Alli90UserAccount;

            }

            public static void WriteExceptions1(IEnumerable<i90UserAccount> exceptionList, string filelocation)
            {


                var newFile = new ExcelPackage(new FileInfo(@filelocation));

                newFile.Workbook.Worksheets.Add("Exceptions");

                ExcelWorksheet exceptions = newFile.Workbook.Worksheets["Exceptions"];


               
                exceptions.Cells[1, 1].Value = "USER";
                exceptions.Cells[1, 2].Value = "TEXT";
                exceptions.Cells[1, 3].Value = "STATUS";
                exceptions.Cells[1, 4].Value = "DATE_CREATION";
                exceptions.Cells[1, 5].Value = "DATE_PERVIOUS_SIGN_ON";
                


                exceptions.Column(1).Width = 25;
                exceptions.Column(2).Width = 25;
                exceptions.Column(3).Width = 30;
                exceptions.Column(4).Width = 30;
                exceptions.Column(5).Width = 30;
               



                exceptions.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                exceptions.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Purple);
                exceptions.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.Purple);
              


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



                int row = 2;
                foreach (var item in exceptionList)
                {
                    exceptions.Cells[row, 1].Value = item.User;
                    exceptions.Cells[row, 2].Value = item.Text;
                    exceptions.Cells[row, 3].Value = item.Status;
                    exceptions.Cells[row, 4].Value = item.Date_Creation;
                    exceptions.Cells[row, 5].Value = item.Date_Previous_Sign_on;



                row++;
                }

                newFile.Save();

            }


        }

    }

