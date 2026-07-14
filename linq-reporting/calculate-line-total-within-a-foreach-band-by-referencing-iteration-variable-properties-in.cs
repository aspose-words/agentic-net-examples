using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report
        string templatePath = "InvoiceTemplate.docx";
        string reportPath = "InvoiceReport.docx";

        // -------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Invoice");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Description]>>");
        builder.Writeln("Qty: <<[item.Quantity]>>");
        builder.Writeln("Price: <<[item.UnitPrice]>>");
        builder.Writeln("Line Total: <<[item.Quantity * item.UnitPrice]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data model
        // -------------------------------------------------
        var model = new ReportModel
        {
            Items = new List<OrderItem>
            {
                new OrderItem { Description = "Widget A", Quantity = 3, UnitPrice = 19.99m },
                new OrderItem { Description = "Widget B", Quantity = 5, UnitPrice = 9.50m },
                new OrderItem { Description = "Widget C", Quantity = 2, UnitPrice = 24.75m }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}

// Root wrapper class referenced in the template as <<[model]>> (named "model" in BuildReport)
public class ReportModel
{
    public List<OrderItem> Items { get; set; } = new();
}

// Data item class used inside the foreach band
public class OrderItem
{
    public string Description { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
