using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV reading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Amount\r\n" +
            "1,Apple,30\r\n" +
            "2,Banana,75\r\n" +
            "3,Cherry,120\r\n" +
            "4,Date,45\r\n" +
            "5,Elderberry,200\r\n");

        // 2. Load CSV and filter rows (Amount > 50).
        var items = new List<Item>();
        var lines = File.ReadAllLines(csvPath);
        for (int i = 1; i < lines.Length; i++) // skip header
        {
            var parts = lines[i].Split(',');
            if (parts.Length != 3) continue;
            if (!double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out double amount))
                continue;
            if (amount <= 50) continue; // filter condition

            items.Add(new Item
            {
                Index = items.Count + 1,
                Name = parts[1],
                Value = amount
            });
        }

        var model = new ReportModel { Items = items };

        // 3. Create the LINQ Reporting template.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Use a numbered list.
        var list = templateDoc.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberArabicDot);
        builder.ListFormat.List = list;

        // Place <<restartNum>> before the foreach in the same numbered paragraph.
        builder.Writeln("<<restartNum>><<foreach [item in Items]>>");

        // Even-indexed items get a light gray background.
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Index]>>. <<[item.Name]>> - <<[item.Value]>> <</backColor>><</if>>");

        // Odd-indexed items have normal formatting.
        builder.Writeln(
            "<<if [item.Index % 2 != 0]>>" +
            "<<[item.Index]>>. <<[item.Name]>> - <<[item.Value]>>" +
            "<</if>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Reset list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // 4. Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model for a CSV row.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public double Value { get; set; }
}

// Wrapper class used as the root data source in the template.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}
