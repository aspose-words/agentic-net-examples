using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Use a conditional expression inside the textColor tag to choose the color
        // based on the numeric Score value.
        // Green  : Score >= 80
        // Orange : 50 <= Score < 80
        // Red    : Score < 50
        builder.Writeln(
            "<<textColor [item.Score >= 80 ? \"Green\" : item.Score >= 50 ? \"Orange\" : \"Red\"]>>" +
            "Item: <<[item.Name]>>, Score: <<[item.Score]>>" +
            " <</textColor>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to a local file.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template back (simulating a separate load step).
        Document doc = new Document(templatePath);

        // Create sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice", Score = 92 },
                new Item { Name = "Bob",   Score = 76 },
                new Item { Name = "Carol", Score = 43 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(reportPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class containing a name and a numeric score.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Score { get; set; }
}
