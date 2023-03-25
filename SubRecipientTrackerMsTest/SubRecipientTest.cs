using ExcelDataReader;
using Microsoft.VisualStudio.TestPlatform.TestHost;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SubrecipientReader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SubRecipientTrackerMsTest
{
    [TestClass]
    public class SubRecipientTest
    {

        [TestMethod]
        public void TestSubrecipientsInSubawardBudgetExample1()
        {

            string dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName;
      
            string excelFilePath = Path.Combine(dir, "files");

       

            // Get the file paths of all Excel spreadsheets in the folder
            string[] filePaths = Directory.GetFiles(excelFilePath, "SubawardBudgetExample1.xlsx");

            // Create a list to store the subrecipient data
            List<string> subrecipientDataList = new List<string>();

            // Loop through each Excel spreadsheet
         
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                // Read the Excel spreadsheet using ExcelDataReader
                using (var stream = File.Open(filePaths[0], FileMode.Open, FileAccess.Read))
                {
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    // Loop through each row in the "G. Other Direct Costs" section
                    while (reader.Read())
                    {
                        if (reader.GetString(1)?.TrimEnd().Contains("Subaward:", StringComparison.OrdinalIgnoreCase) == true)
                        {
                        subrecipientDataList.Add(reader.GetString(2));

                        }
                    }
                }
            
                            // Arrange
                            var expectedSubrecipients = new string[] { "Indiana", "Mayo", "Purdue", "Florida" };
          

            // Assert
            Assert.AreEqual(expectedSubrecipients.Length, subrecipientDataList.Count);
            foreach (var expectedSubrecipient in expectedSubrecipients)
            {
                Assert.IsTrue(subrecipientDataList.Any(sd => sd == expectedSubrecipient));
            }
        }
}
}
