using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model with a single existing object.
    public class Model
    {
        public Existing ExistingObject { get; set; } = new Existing();
    }

    public class Existing
    {
        public string Value { get; set; } = "Present";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and add LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Tag referencing a missing object – will be treated as null.
            builder.Writeln("Missing field output: <<[MissingObject.Name]>>");

            // Tag referencing an existing field – will be populated normally.
            builder.Writeln("Existing field output: <<[model.ExistingObject.Value]>>");

            // Prepare the data source.
            Model model = new Model();

            // Configure the reporting engine to allow missing members and embed error messages.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages,
                MissingMemberMessage = "NULL"
            };

            // Build the report. The overload with the data source name allows the template to reference the root object.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated document.
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // Output the result of the build operation.
            Console.WriteLine($"Report build successful: {success}");
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
