using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const string templatePath = "Template.docx";
        const string pdfPath = "Report.pdf";

        // -------------------------------------------------
        // 1. Create the Word template with LINQ Reporting tags
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header showing a property from the root model
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln();

        // Begin foreach loop over Items collection
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table inside the foreach block
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row for each item
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // End the table
        builder.EndTable();

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template (ensures it is fully persisted)
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample data model
        // -------------------------------------------------
        ReportModel model = new()
        {
            CustomerName = "Acme Corp",
            Items = new()
            {
                new Item { Index = 1, Name = "Widget" },
                new Item { Index = 2, Name = "Gadget" },
                new Item { Index = 3, Name = "Doohickey" }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using LINQ Reporting engine
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 5. Export the rendered document to PDF
        // -------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}

// -----------------------------------------------------------------
// Public data model classes (must be public with public properties)
// -----------------------------------------------------------------
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
