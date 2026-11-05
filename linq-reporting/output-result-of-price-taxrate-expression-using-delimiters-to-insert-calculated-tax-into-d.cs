using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Invoice
{
    public decimal Price { get; set; } = 100m;
    public decimal TaxRate { get; set; } = 0.07m;
}

public class Program
{
    public static void Main()
    {
        // Create template document with LINQ Reporting tags.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Price: <<[model.Price]>>");
        builder.Writeln("Tax Rate: <<[model.TaxRate]>>");
        builder.Writeln("Tax Amount: <<[model.Price * model.TaxRate]>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare data model.
        var invoice = new Invoice();

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, invoice, "model");

        // Save the generated document.
        var outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
