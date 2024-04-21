using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data.Common;


namespace IEEEElectHelper;

public class Loader
{
    public static string debugpath = "assets\\results\\debug.txt";
    public static string path;

    public static void Testagain()
    {
        bool ismember = true;
        string excelFilePath = @"C:\Users\xtige\RiderProjects\IEEEElectHelper\IEEEElectHelper\bin\Debug\net8.0\assets\test.xlsx";
        if (!File.Exists(debugpath))
        {
            File.Create(debugpath).Dispose();
        }
        File.WriteAllText(debugpath, "");
        File.WriteAllText("assets\\results\\skipped\\skipped.txt", "");

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Réponses au formulaire 1"];
            int rowCount = worksheet.Dimension.Rows;

            Dictionary<string, Dictionary<string, int>> chapterResults = new Dictionary<string, Dictionary<string, int>>(); 

            for (int row = 2; row <= rowCount; row++)
            {
                string ieeeID = worksheet.Cells[row, 2].Value.ToString(); 
                for (int col = 3; col <= worksheet.Dimension.Columns; col++) 
                {
                    string position = worksheet.Cells[1, col].Value.ToString();
                    if (position.Contains("Are"))
                    {
                        continue;
                    }

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

                        if (Functions.IsMemberOfChapter(ieeeID, position))
                        {
                            chapterResults[position][vote]++;
                            string name = Loader.returnname(ieeeID);
                            string chapter = Functions.GetChapterName(position);
                            string ph = $"{ieeeID} ({name}) voted for {vote} {position} \n";
                            File.AppendAllText("assets\\results\\debug.txt", ph );

                        }
                        else
                        {
                            if (!string.IsNullOrWhiteSpace(ieeeID))
                            Writer.writeskipped(position, ieeeID, vote);
                        }

                    }
                }
                File.AppendAllText("assets\\results\\debug.txt", "\n");
            }



            foreach (var kvp in chapterResults)
            {

                string position = kvp.Key;
                if (position.ToLower().Contains("are"))
                {
                    continue;
                }
                Dictionary<string, int> optionResults = kvp.Value;

                // Write to text file
                string fileName = $"assets\\results\\{position}_Results.txt";
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
    public static string returnname(string id)
    {
        using (var package = new ExcelPackage(new FileInfo("assets\\mem.xlsx")))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.End.Row;

            for (int row = 1; row <= rowCount; row++)
            {
                // Read the ID from the first column
                if (worksheet.Cells[row, 2].Value != null)
                {
                    string currentId = worksheet.Cells[row, 2].Value.ToString();
                    // Check if it matches the searchId
                    if (currentId.Equals(id, StringComparison.OrdinalIgnoreCase))
                    {
                        // Return the name found in the second column
                        string firstname = worksheet.Cells[row, 4].Value.ToString();
                        string secondname = worksheet.Cells[row, 3].Value.ToString();
                        string fullname = firstname + " " + secondname;
                        return fullname;
                    }
                }
            }
        }
        return null;
    }

}