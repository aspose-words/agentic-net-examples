using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and result documents
        const string templatePath = "Template.docx";
        const string resultPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create a DOCX template with LINQ Reporting tags
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title
        builder.Writeln("Sample Report");
        builder.Writeln();

        // Opening tag for the collection (foreach)
        builder.Writeln("<<foreach [item in Items]>>");
        // Content for each item
        builder.Writeln("Item: <<[item.Name]>>  Value: <<[item.Value]>>");
        // Closing tag for the collection
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template (ensures BuildReport is called after loading)
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample data model
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Value = 10 },
                new Item { Name = "Banana", Value = 20 },
                new Item { Name = "Cherry", Value = 30 }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using ReportingEngine
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options needed for this simple example
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report
        // -------------------------------------------------
        doc.Save(resultPath);
    }
}

// -------------------------------------------------
// Data model classes (public, non‑nullable properties initialized)
// -------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
