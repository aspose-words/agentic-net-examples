using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Static paragraph
        builder.Writeln("Customer Report");

        // Paragraph that contains only a LINQ Reporting tag (will be removed)
        builder.Writeln("<<[model.Name]>>");

        // Paragraph with tag inside text (will remain)
        builder.Writeln("Name: <<[model.Name]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Sample data model
        ReportModel model = new ReportModel { Name = "John Doe" };

        // Configure the reporting engine to remove empty paragraphs
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report
        engine.BuildReport(doc, model, "model");

        // Save the final report
        doc.Save(outputPath);
    }

    public class ReportModel
    {
        public string Name { get; set; } = "";
    }
}
