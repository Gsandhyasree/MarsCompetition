using System;
using System.Xml.Linq;
using MarsCompetition.Pages;

using NUnit.Framework;
using static System.Net.Mime.MediaTypeNames;
using static MarsCompetition.Global.GlobalDefinitions;
using static MarsCompetition.Pages.ShareSkill;

namespace MarsCompetition
{
    [TestFixture]
    [Parallelizable]
    internal class Tests : Global.Base
    {

        ManageListings manageListingObj;
        ShareSkill shareSkillObj;

        [Test, Order(1)]
        public void TC1a_WhenIEnterListing()
        {
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            ManageListings manageListingObj = new ManageListings();
            manageListingObj.AddListing(2, "ManageListings");

        }
        [Test, Order(2)]
        public void TC1b_ThenListingIsCreated()
        {
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            manageListingObj = new ManageListings();
            VerifyListingDetails(2, "ManageListings");
        }

        [Test, Order(3)]
        public void TC2a_WhenIEditListing()
        {
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            manageListingObj = new ManageListings();
            manageListingObj.EditListing(2, 3, "ManageListings");
        }

        [Test, Order(4)]
        public void TC2b_ThenListingIsEdited()
        {
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            VerifyListingDetails(3, "ManageListings");
        }

        [Test, Order(5)]
        public void TC3a_WhenIDeleteListing()
        {
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            manageListingObj = new ManageListings();
            manageListingObj.DeleteListing(3, "ManageListings");
        }

        [Test, Order(6)]
        public void TC3b_ThenListingIsDeleted()
        {
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            VerifyDelete(3, "ManageListings");
        }


        public void VerifyListingDetails(int rowNumber, string worksheet)
        {
            //Click on view Listing
            manageListingObj = new ManageListings();
            shareSkillObj = new ShareSkill();
            manageListingObj.ViewListing(rowNumber, worksheet);

            Listing excel = new Listing();
            Listing web = new Listing();


            shareSkillObj.GetExcel(rowNumber, worksheet, out excel);

            shareSkillObj.GetWeb(out web);

            //Assertions
            Assert.Multiple(() =>
            {

                //Verify expected Title vs actual Title

                Console.WriteLine("Excel   |" +  excel.title + "| Web    |"  + web.title+ "| ");
                Assert.AreEqual(excel.title, web.title);

                //Verify expected Description vs actual Description
                Assert.AreEqual(excel.description, web.description);

                //Verify expected Category vs actual Category
                Assert.AreEqual(excel.category, web.category);

                //Verify expected Subcategory vs actual Subcategory
                Assert.AreEqual(excel.subcategory, web.subcategory);

                //Verify expected ServiceType vs actual ServiceType
                string serviceTypeText = "Hourly";
                if (excel.serviceType == "One-off service")
                    serviceTypeText = "One-off";
                Assert.AreEqual(serviceTypeText, web.serviceType);

                //Verify expected StartDate vs actual StartDate
                string expectedStartDate = DateTime.Parse(excel.startDate).ToString("yyyy-MM-dd");
                Assert.AreEqual(expectedStartDate, web.startDate);

                //Verify expected EndDate vs actual EndDate
                string expectedEndDate = DateTime.Parse(excel.endDate).ToString("yyyy-MM-dd");
                Assert.AreEqual(expectedEndDate, web.endDate);

                //Verify expected LocationType vs actual LocationType
                string expectedLocationType = excel.locationType;
                if (expectedLocationType.Equals("On-site"))
                    expectedLocationType = "On-Site";
                Assert.AreEqual(expectedLocationType, web.locationType);

                //Verify Skills Trade
                if (excel.skillTrade == "Credit")
                    Assert.AreEqual("None Specified", shareSkillObj.GetSkillTrade("Credit"));
                else
                    Assert.AreEqual(excel.skillExchange, shareSkillObj.GetSkillTrade("Skill-exchange"));
            });

        }

        public void VerifyDelete(int rowNumber, string worksheet)
        {
            manageListingObj = new ManageListings();
            //Populate excel data
            ExcelLib.PopulateInCollection(ExcelPath, worksheet);
            string title = ExcelLib.ReadData(rowNumber, "Title");

            //Click on Manage Listing
            manageListingObj.GoToManageListings();

            //Assertion
            Assert.AreNotEqual(title, manageListingObj.FindTitle(title), "Delete Failed");
        }

      
    }
}

