using Newtonsoft.Json;
using OfficeOpenXml;

namespace IEEEElectHelper;

public class Loader
{
    public static void Test()
    {
        string excelFilePath = @"C:\Users\xtige\RiderProjects\IEEEElectHelper\IEEEElectHelper\bin\Debug\net8.0\assets\Members.xlsx";

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
            int rowCount = worksheet.Dimension.Rows;

            var chapters = new Dictionary<string, List<string>>();
            string currentChapter = null;

            for (int row = 1; row <= rowCount; row++)
            {
                string cellValue = worksheet.Cells[row, 1].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    currentChapter = cellValue;
                    if (!chapters.ContainsKey(currentChapter))
                    {
                        chapters.Add(currentChapter, new List<string>());
                    }
                }
                else if (!string.IsNullOrWhiteSpace(currentChapter))
                {
                    string ieeeId = worksheet.Cells[row, 2].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(ieeeId))
                    {
                        chapters[currentChapter].Add(ieeeId);
                    }
                }
            }
            foreach (var chapter in chapters)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Chapter: {chapter.Key}");
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"Members: {string.Join(", ", chapter.Value)}");
            }

            string jsonFilePath = "assets\\output.json";
            File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(chapters, Formatting.Indented));

            Console.WriteLine("Parsing completed. Output saved to output.json");
        }

    }



    public static void Testagain()
    {
        bool ismember = true;
        string excelFilePath = @"C:\Users\xtige\RiderProjects\IEEEElectHelper\IEEEElectHelper\bin\Debug\net8.0\assets\tt.xlsx";

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
            int rowCount = worksheet.Dimension.Rows;



            Dictionary<string, Dictionary<string, int>> chapterResults = new Dictionary<string, Dictionary<string, int>>(); 

            for (int row = 2; row <= rowCount; row++)
            {
                string ieeeID = worksheet.Cells[row, 4].Value.ToString(); 


                for (int col = 6; col <= worksheet.Dimension.Columns; col++) 
                {
                    string position = worksheet.Cells[1, col].Value.ToString(); 
                    string vote = worksheet.Cells[row, col].Value?.ToString(); 

                    if (!string.IsNullOrEmpty(vote))
                    {
                        if (!chapterResults.ContainsKey(position))
                        {
                            chapterResults[position] = new Dictionary<string, int>(); 
                        }

                        if (!chapterResults[position].ContainsKey(vote))
                        {
                            chapterResults[position][vote] = 0; 
                        }

                        if (ismember)
                        {
                            chapterResults[position][vote]++;
                        }
                    }
                }
            }
            foreach (var kvp in chapterResults)
            {
                string position = kvp.Key;
                Dictionary<string, int> optionResults = kvp.Value;

                // Write to text file
                string fileName = $"assets\\{position}_Results.txt";
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileName))
                {
                    file.WriteLine($"Position: {position}");
                    foreach (var optionKvp in optionResults)
                    {
                        string option = optionKvp.Key;
                        int voteCount = optionKvp.Value;
                        file.WriteLine($"{option}: {voteCount}");
                    }
                }
            }

        }
    }

}