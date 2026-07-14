using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "Aspose User";
    }

    public class Program
    {
        public static void Main()
        {
            // Folder for temporary files.
            const string outputDir = "Output";
            System.IO.Directory.CreateDirectory(outputDir);

            // 1. Create a template document with a LINQ Reporting tag.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Hello, <<[model.Name]>>!");
            string templatePath = System.IO.Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // 2. Load the template (simulating a separate load step).
            Document doc = new Document(templatePath);

            // 3. Restrict access to sensitive .NET types before any report generation.
            // Example: prevent the template from accessing System.Environment and System.IO.FileInfo.
            ReportingEngine.SetRestrictedTypes(typeof(System.Environment), typeof(System.IO.FileInfo));

            // 4. Prepare the data source.
            Model model = new Model();

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string resultPath = System.IO.Path.Combine(outputDir, "Result.docx");
            doc.Save(resultPath);
        }
    }
}
