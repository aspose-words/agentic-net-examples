using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with nullable decimal fields.
        var model = new ReportModel
        {
            Discount = 12.5m,
            Tax = null // Tax is missing for this example.
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Generates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Discount: <<[model.Discount]>>");
        builder.Writeln("Tax: <<[model.Tax]>>");
        builder.Writeln("Combined (Discount + Tax): <<[model.Combined]>>");

        doc.Save(filePath);
    }
}

// Data model used by the report. All members are public.
public class ReportModel
{
    // Nullable decimal fields.
    public decimal? Discount { get; set; }
    public decimal? Tax { get; set; }

    // Combined value using the lifted addition operator.
    // The result is null if either operand is null.
    public decimal? Combined => Discount + Tax;
}
