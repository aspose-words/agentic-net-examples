using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model required by the reporting engine.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Title { get; set; } = string.Empty;
    }

    // Wrapper class used to provide an optional document for inclusion.
    public class IncludeSource
    {
        // If the file exists this property holds the loaded document; otherwise it remains null.
        public Document? Document { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a blank document and add LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("=== Report Start ===");

            // Use the <<doc>> tag to include another document.
            // The IncludeSource.Document property will be null if the file does not exist,
            // causing the engine to skip the inclusion without throwing an error.
            builder.Writeln("<<doc [src.Document]>>");

            builder.Writeln("=== Report End ===");

            // Prepare a simple root object for the report.
            ReportModel model = new ReportModel { Title = "Sample Report" };

            // Prepare the optional include source.
            IncludeSource src = new IncludeSource();
            string includePath = Path.Combine(Environment.CurrentDirectory, "missing_file.txt");
            if (File.Exists(includePath))
            {
                // Load the file as a Word document if it exists.
                src.Document = new Document(includePath);
            }
            else
            {
                // No file – keep Document null so the engine treats it as optional.
                src.Document = null;
            }

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members (including a null Document) to be treated as empty.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report. The root object name must match the name used in the template.
            // Pass both the main model and the include source.
            engine.BuildReport(doc, new object[] { model, src }, new[] { "model", "src" });

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
