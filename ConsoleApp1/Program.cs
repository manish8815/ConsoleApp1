using ClosedXML.Excel;
using HtmlAgilityPack;

class Program
{
    static void Main()
    {
        string path = "D:/Dotnet/ConsoleApp1/ConsoleApp1/data/klg24-25TRIALBALANCE.htm";

        try
        {
            ReadHtml(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading file: {ex.Message}");
        }
    }

    static void ReadHtml(string path)
    {
        var htmlDoc = new HtmlDocument();
        htmlDoc.Load(path);
        List<string> lines = new List<string>();
        var groupNodes = htmlDoc.DocumentNode.SelectNodes("//span[@class='CPI_12_Courier' and contains(., 'Group :')]");

        if (groupNodes != null)
        {
            foreach (var groupNode in groupNodes)
            {
                var boldText = groupNode.SelectSingleNode(".//b");
                if (boldText != null)
                {
                    string groupName = boldText.InnerText.Trim();
                    Console.WriteLine(groupName);
                    lines.Add(groupName);
                }
            }
            CreateExcelSheet(lines);
        }
        else
        {
            Console.WriteLine("No group headers found in the document.");
        }
    }

    static void CreateExcelSheet(List<string> lines)
    {
        string filePath = "D:/xltest.xlsx";

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Busy_Soft");

            // Add header information
            worksheet.Cell("B1").Value = "Busy Software Group Name Fetched..";
            var headerCell = worksheet.Range("B1");
            headerCell.Style.Font.Bold = true;
            worksheet.Column(2).Width = 50;

            // Set column headers
            worksheet.Cell("A2").Value = "S.No.";
            worksheet.Cell("B2").Value = "Group Name";
            worksheet.Range("A2:B2").Style.Font.Bold = true;

            int row = 3;
            for (int i = 0; i < lines.Count; i++)
            {
                worksheet.Cell(row, 1).Value = i + 1;
                worksheet.Cell(row, 2).Value = lines[i];
                row++;
            }

            // Save the workbook
            workbook.SaveAs(filePath);
            Console.WriteLine($"Excel file saved to: {filePath}");
        }
    }
}