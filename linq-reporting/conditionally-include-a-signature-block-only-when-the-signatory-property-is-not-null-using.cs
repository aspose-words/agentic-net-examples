using System;
using System.Text;
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
        // Register code page provider for full encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Step 1: Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Add a title placeholder.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // Conditional block: include the signature line only when Signatory is not null.
        builder.Writeln("<<if [model.Signatory != null]>>Signature: <<[model.Signatory]>> <</if>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            Title = "Monthly Sales Report",
            Signatory = "John Doe"
            // To test the absence of a signature block, set Signatory = null;
        };

        // Step 3: Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);
    }
}
