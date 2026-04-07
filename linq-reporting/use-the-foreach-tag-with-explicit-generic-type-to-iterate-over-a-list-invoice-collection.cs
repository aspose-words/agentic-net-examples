using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Invoice
{
    public int Number { get; set; }
    public decimal Amount { get; set; }
}

public class ReportModel
{
    // Collection of invoices to be used as the data source.
    public List<Invoice> Invoices { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Prepare sample data.
        var model = new ReportModel();
        model.Invoices.Add(new Invoice { Number = 1001, Amount = 250.75m });
        model.Invoices.Add(new Invoice { Number = 1002, Amount = 489.00m });
        model.Invoices.Add(new Invoice { Number = 1003, Amount = 123.45m });

        // 2. Create a Word document template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a foreach tag (correct syntax) to iterate over the Invoices collection.
        builder.Writeln("<<foreach [inv in Invoices]>>");
        builder.Writeln("Invoice #: <<[inv.Number]>> - Amount: $<<[inv.Amount]>>");
        builder.Writeln("<</foreach>>");

        // 3. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 4. Save the generated document.
        doc.Save("InvoiceReport.docx");
    }
}
