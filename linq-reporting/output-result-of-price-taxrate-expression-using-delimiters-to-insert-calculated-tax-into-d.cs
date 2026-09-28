using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Invoice
{
    public decimal Price { get; set; } = 0m;
    public decimal TaxRate { get; set; } = 0m;
}

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Paths for template and result documents.
        string templatePath = Path.Combine(workDir, "Template.docx");
        string resultPath = Path.Combine(workDir, "Result.docx");

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Price: <<[Price]>>");
        builder.Writeln("Tax Rate: <<[TaxRate]>>");
        builder.Writeln("Tax Amount: <<[Price * TaxRate]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare sample data.
        Invoice model = new Invoice
        {
            Price = 123.45m,
            TaxRate = 0.07m // 7% tax
        };

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(resultPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
