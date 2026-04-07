using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    // Name of the item – initialized to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    // Collection of items – initialized with an empty list.
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        const string templateFile = "Template.docx";

        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The template uses ElementAt to fetch the fifth element (index 4) without a loop.
        builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");

        // Save the template to disk.
        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 2. Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel();

        // Populate the collection with sample data (more than five items).
        for (int i = 1; i <= 10; i++)
        {
            model.Items.Add(new Item { Name = $"Item {i}" });
        }

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templateFile);

        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputFile = "Report.docx";
        reportDoc.Save(outputFile);

        // Indicate successful completion (no interactive input required).
        Console.WriteLine($"Report generated: {outputFile}");
    }
}
