using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a simple template with a missing member reference.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Name: <<[model.Name]>>");
        builder.Writeln("Missing Property: <<[model.MissingProperty]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            Name = "John Doe"
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Enable inline error messages so missing members are shown in the output.
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns true if the report was built without fatal errors.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Public data model class.
    public class ReportModel
    {
        // Initialized to avoid nullable warnings.
        public string Name { get; set; } = "";
    }
}
