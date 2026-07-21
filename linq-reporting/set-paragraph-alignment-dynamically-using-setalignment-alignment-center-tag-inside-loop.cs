using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some data sources).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // 1. Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a LINQ Reporting foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");
        // Insert the item text. The unsupported <<setAlignment>> tag has been removed.
        builder.Writeln("<<[item.Text]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template document for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Text = "First centered paragraph." },
                new Item { Text = "Second centered paragraph." },
                new Item { Text = "Third centered paragraph." }
            }
        };

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Text { get; set; } = string.Empty;
}
