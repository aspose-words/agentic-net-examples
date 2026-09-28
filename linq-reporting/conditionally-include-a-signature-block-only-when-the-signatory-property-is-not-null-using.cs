using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = "";
    public string? Signatory { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create the template document
        var templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title placeholder
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Conditional signature block (included only when Signatory is not null)
        builder.Writeln("<<if [model.Signatory != null]>>");
        builder.Writeln("Signed by: <<[model.Signatory]>>");
        builder.Writeln("<</if>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Prepare sample data
        var model = new ReportModel
        {
            Title = "Sample Report",
            Signatory = "John Doe" // Change to null to omit the signature block
        };

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
