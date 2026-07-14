using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the template and final PDF.
        const string templatePath = "ReportTemplate.docx";
        const string pdfPath = "Report.pdf";

        // 1. Create the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header with a simple field.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Loop over the items collection.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template back (required before building the report).
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare sample data.
        Order sampleOrder = new Order
        {
            CustomerName = "Acme Corp",
            Items =
            {
                new Item { Index = 1, Name = "Widget" },
                new Item { Index = 2, Name = "Gadget" },
                new Item { Index = 3, Name = "Doohickey" }
            }
        };

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, sampleOrder, "order");

        // 5. Export the rendered document to PDF.
        loadedTemplate.Save(pdfPath, SaveFormat.Pdf);
    }
}

// Root data model.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

// Item model used inside the collection.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
