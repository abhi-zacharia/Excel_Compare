using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using System.Collections.Generic;
using System.Linq;


namespace StrataUsers
{
    [TestFixture]
    public class LeanFtTest : UnitTestClassBase
    {
        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            // Before each test
        }

        [Test]
        public void Strata_User()
        {

            List<CurrentUsers> AllCurrentUsersList = CurrentUserList_ReadExcel.ReadExcelForCurrentUsers("Current User List", @"C:\Automation\Judy_Data\Strata users and HR Employees April 18.xlsx", 1);

            List<HREmployee> AllHREmployee = HRList_ReadExel.ReadExcelForHREmployeeList("HR Employee List", @"C:\Automation\Judy_Data\Strata users and HR Employees April 18.xlsx", 2, 3);

            //AllCurrentUsers.Sort();
            // AllHREmployee.Sort();

            List<string> Current = new List<string>();
            List<string> Existing = new List<string>();

            foreach (var item in AllCurrentUsersList)
            {
                Current.Add(item.Name.Trim());
            }

            foreach (var item in AllHREmployee)
            {
                Existing.Add(item.FirstName.Trim());
            }

            List<string> nomatch = Current.Except(Existing).ToList();


            CurrentUserList_ReadExcel.WriteExceptions(nomatch, @"C:\Automation\Judy_Data\Strata users and HR Employees April 18.xlsx");



        }

        [Test]

        public void CDlClassic_User()
        {
            List<CDLClassicUsers> AllCDLUsers = CDLClassicUsersList.ReadExcelForCDLUsers("UserList", @"C:\Automation\Judy_Data\CDL Classic users and HR Employees April 18.xlsx",1,2,3,4);

            List<HREmployee> AllHREmployee = HRList_ReadExel.ReadExcelForHREmployeeList("HR List", @"C:\Automation\Judy_Data\CDL Classic users and HR Employees April 18.xlsx", 2, 3);

            var nomatch = AllCDLUsers.Where(p => !AllHREmployee.Any(p2 => string.Format("{0} {1}", p2.FirstName.Trim(), p2.Surname.Trim()) == p.Name.Trim())).Where(p => p.Name != null | p.Name != "").ToList();

            CDLClassicUsersList.WriteExceptions(nomatch, @"C:\Automation\Judy_Data\CDL Classic users and HR Employees April 18.xlsx");


        }
        [Test]
        public void OGI_User()
        {

            List<OGIUser> AllOGIUsers = OGIUserList_ReadExcel.ReadExcelForOGIUsers("OGI Users", @"C:\Automation\Judy_Data\OGI User Review Nov 19.xlsx", 1,2,3,4,5,6);

            List<HREmployee> AllHREmployee = HRList_ReadExel.ReadExcelForHREmployeeList1("Employee List", @"C:\Automation\Judy_Data\OGI User Review Nov 19.xlsx", 2, 3);
          
            var nomatch = AllOGIUsers.Where(p=> !AllHREmployee.Any(p2=> string.Format("{0} {1}", p2.FirstName.Trim(), p2.Surname.Trim()) == p.Name.Trim())).Where(p=> p.Name != null | p.Name != "").ToList();
            
            OGIUserList_ReadExcel.WriteExceptions(nomatch, @"C:\Automation\Judy_Data\OGI User Review Nov 19.xlsx");

            
        }

        [Test]

        public void SOX_Leavers()

        {

            List<SOXLeavers> ALLSOAllSOXLeavers = SOXLeaversFullListReadExcel.ReadExcelForSoXLeavers("i90 Full list - May 18", @"C:\Automation\Judy_Data\SOX_Leavers_full_list_encrypted.xlsx", 1, 2, 3, 4, 5, 6, 7, 8);

            List<HREmployee> AllHREmployee = HRList_ReadExel.ReadExcelForHREmployeeList1("HR Staff List - May 18", @"C:\Automation\Judy_Data\SOX_Leavers_full_list_encrypted.xlsx", 2, 3);                     
           
            var nomatch = ALLSOAllSOXLeavers.Where(p => !AllHREmployee.Any(p2 => string.Format("{0} {1}", p2.FirstName.Trim(), p2.Surname.Trim()) == p.Text.Trim())).Where(p => p.Text != null | p.Text != "").ToList();
           
            SOXLeaversFullListReadExcel.WriteExceptions1(nomatch, @"C:\Automation\Judy_Data\SOX_Leavers_full_list_encrypted.xlsx");
        }

        [Test]

        public void i90User_Accounts()
        {
            List<i90UserAccount> Alli90UserAccounts = I90UserAccountList_ReadExcel.ReadExcelFori90Users("i90 User Accounts", @"C:\Automation\Judy_Data\Acitvei90 Users and Current HR Employees Aug 18.xlsx",1,2,3,4,5);

            List<HREmployee> AllHREmployee = HRList_ReadExel.ReadExcelForHREmployeeList1("HR Employee Record", @"C:\Automation\Judy_Data\Acitvei90 Users and Current HR Employees Aug 18.xlsx",2, 3);

            var nomatch = Alli90UserAccounts.Where(p => !AllHREmployee.Any(p2 => string.Format("{0} {1}", p2.FirstName.Trim(), p2.Surname.Trim()) == p.Text.Trim())).Where(p => p.Text != null | p.Text != "").ToList();

            I90UserAccountList_ReadExcel.WriteExceptions1(nomatch, @"C:\Automation\Judy_Data\Acitvei90 Users and Current HR Employees Aug 18.xlsx");
        }


        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }

}