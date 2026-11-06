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

        // Paths for the template and the generated document
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a template document that contains a missing member tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Name: <<[model.Name]>>");
        builder.Writeln("Missing field: <<[model.NonExisting]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Sample data model (does NOT contain the missing member)
        SampleModel model = new SampleModel { Name = "John Doe" };

        // Configure the reporting engine
        ReportingEngine engine = new ReportingEngine();
        // Enable handling of missing members without throwing an exception
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report
        bool success = engine.BuildReport(doc, model, "model");

        // The ReportingEngine in the current version does not expose a MissingMemberLog property.
        // When AllowMissingMembers is enabled, missing members are simply ignored.
        // For demonstration, we output a message indicating whether the report was generated.
        if (success)
        {
            Console.WriteLine("Report generated successfully. Missing members (if any) were ignored.");
        }
        else
        {
            Console.WriteLine("Report generation failed.");
        }

        // Save the generated document
        doc.Save(outputPath);
    }
}

// Public data model class
public class SampleModel
{
    public string Name { get; set; } = "";
}
