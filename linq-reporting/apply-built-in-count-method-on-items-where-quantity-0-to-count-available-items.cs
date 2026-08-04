using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a tag that counts items with Quantity > 0.
        // The tag uses the built‑in Count method with a predicate.
        builder.Writeln("Available items count: <<[model.Items.Count(i => i.Quantity > 0)]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Prepare the data ----------
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Quantity = 5 },
                new Item { Name = "Banana", Quantity = 0 },
                new Item { Name = "Cherry", Quantity = 12 },
                new Item { Name = "Date",   Quantity = -3 } // Negative quantity is treated as unavailable.
            }
        };

        // ---------- Load the template and build the report ----------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model; the root name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
