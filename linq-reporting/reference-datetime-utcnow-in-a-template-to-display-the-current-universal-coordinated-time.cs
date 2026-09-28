using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words in some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create a template document that contains a LINQ Reporting tag.
        // The tag references a property on the root model (CurrentUtc).
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Current UTC time: <<[model.CurrentUtc]>>");

        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using an empty model that provides the CurrentUtc property.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        var model = new Model(); // root object
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Model class exposed to the reporting engine.
    public class Model
    {
        // Returns the current UTC time when the report is generated.
        public DateTime CurrentUtc => DateTime.UtcNow;
    }
}
