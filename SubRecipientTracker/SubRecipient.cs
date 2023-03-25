using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using ExcelDataReader;

namespace SubrecipientReader
{
  public  class SubRecipient
    {
        static void Main(string[] args)
        {
            
            string dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName;
            // Read the file as one string. 
    
            string excelFilePath = Path.Combine(dir, "files");

   

            // Get the file paths of all Excel spreadsheets in the folder
            string[] filePaths = Directory.GetFiles(excelFilePath, "*.xlsx");

            // Create a list to store the subrecipient data
            List<SubrecipientData> subrecipientDataList = new List<SubrecipientData>();

            // Loop through each Excel spreadsheet
            foreach (string filePath in filePaths)
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                // Read the Excel spreadsheet using ExcelDataReader
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    // Loop through each row in the "G. Other Direct Costs" section
                    while (reader.Read())
                    {
                        if (reader.GetString(1)?.TrimEnd().Contains("Subaward:", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            // Extract the subrecipient name and subaward amount

                            string subrecipientName = reader.GetString(2)?.Trim();

                            // Extract the subrecipient name and subaward amount
                            if (reader.GetString(1)?.TrimEnd().Length != "Subaward:".Length)
                            {
                                string[]? subrecipientNames = reader.GetString(1)?.TrimEnd().Split(":");
                                subrecipientName= !string.IsNullOrWhiteSpace(subrecipientNames[1]) ? subrecipientNames[1] : "N/A";

                            }
                            subrecipientName = subrecipientName == null ?"N/A":subrecipientName ;


                            int lastColumnIndex = reader.FieldCount - 1;
                            double subawardAmount = 0;

                            // Loop backwards through the row to find the last non-empty cell
                            for (int i = lastColumnIndex; i > 1; i--)
                            {
                                if (!reader.IsDBNull(i) && !string.IsNullOrWhiteSpace(reader.GetValue(i).ToString()))
                                {
                                    if (reader.GetDouble(i)==0)
                                    {
                                        subawardAmount = reader.GetDouble(i-1);
                                        break;
                                    }
                                    subawardAmount = reader.GetDouble(i);
                                    break;
                                }
                            }

                            // Add the subrecipient data to the list
                            subrecipientDataList.Add(new SubrecipientData { FilePath = filePath, SubrecipientName = subrecipientName?.Trim(), SubawardAmount = subawardAmount });
                        }
                    }
                }
            }

            // Output the subrecipient data to the console
            foreach (SubrecipientData subrecipientData in subrecipientDataList)
            {
                Console.WriteLine($"{Path.GetFileName(subrecipientData.FilePath)}\t{subrecipientData.SubrecipientName}\t{subrecipientData.SubawardAmount}");
            }

            // Output the total subaward amount for each subrecipient
            var subrecipientTotals = subrecipientDataList
                .GroupBy(sd => sd.SubrecipientName)
                .Select(g => new { SubrecipientName = g.Key, TotalSubawardAmount = g.Sum(sd => sd.SubawardAmount) })
                .OrderByDescending(g => g.TotalSubawardAmount);

            Console.WriteLine("\nSubrecipient Totals:\n");

            foreach (var subrecipientTotal in subrecipientTotals)
            {
                Console.WriteLine($"{subrecipientTotal.SubrecipientName}\t{subrecipientTotal.TotalSubawardAmount}");
            }
        }
    }

  public  class SubrecipientData
    {
        public string FilePath { get; set; }
        public string SubrecipientName { get; set; }
        public double SubawardAmount { get; set; }
    }
}

