using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public int Stock { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product stock status:");
        // Conditional block: show "In stock" only when Stock > 0.
        builder.Writeln("<<if [model.Stock > 0]>>In stock<</if>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Sample data.
        var model = new Product { Stock = 5 };

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
