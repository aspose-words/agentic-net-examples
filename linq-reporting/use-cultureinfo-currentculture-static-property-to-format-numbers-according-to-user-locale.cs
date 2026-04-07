using System;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data model.
        var model = new ReportModel
        {
            Description = "Sample amount",
            Amount = 12345.67m
        };

        // Create the template document.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report");
        builder.Writeln("Description: <<[model.Description]>>");
        builder.Writeln("Amount (formatted): <<[model.FormattedAmount]>>");

        doc.Save(path);
    }
}

public class ReportModel
{
    public string Description { get; set; } = string.Empty;
    public decimal Amount { get; set; }

    // Returns the amount formatted according to the current culture.
    public string FormattedAmount => Amount.ToString("N", CultureInfo.CurrentCulture);
}
