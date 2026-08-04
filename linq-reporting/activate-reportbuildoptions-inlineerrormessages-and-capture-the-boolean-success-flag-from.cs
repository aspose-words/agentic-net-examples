using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a LINQ Reporting tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<[model.Name]>>!");

        // Prepare the data model.
        var model = new ReportModel { Name = "World" };

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("ReportOutput.docx");

        // Output the success flag.
        Console.WriteLine($"BuildReport success: {success}");
    }

    // Simple data model class used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "";
    }
}
