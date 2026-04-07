using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // This value is not numeric; using it in a numeric expression will cause a conversion error.
        public string Age { get; set; } = "abc";

        // This value is not a valid date; accessing date members will cause a conversion error.
        public string BirthDate { get; set; } = "not-a-date";

        // Regular string property that will be displayed correctly.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Create a new blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading.
            builder.Writeln("=== LINQ Reporting with Inline Error Messages ===");
            builder.Writeln();

            // Insert a correct expression.
            builder.Writeln("Name: <<[model.Name]>>");

            // Insert expressions that will fail during evaluation.
            // The division expects a numeric value, but Age is a non‑numeric string.
            builder.Writeln("Age divided by 2 (will cause error): <<[model.Age] / 2>>");

            // Attempt to access the Year property of a date, but BirthDate is not a valid date string.
            builder.Writeln("Birth year (will cause error): <<[model.BirthDate].Year>>");

            // Save the template to a temporary file (optional, can be omitted).
            const string templatePath = "TemplateWithErrors.docx";
            doc.Save(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The third argument is the name used in the template to reference the root object.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the success flag to the console.
            Console.WriteLine($"Report build success flag: {success}");

            // Save the generated report.
            const string outputPath = "ReportWithInlineErrors.docx";
            doc.Save(outputPath);

            // Inform the user where the files are saved.
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
