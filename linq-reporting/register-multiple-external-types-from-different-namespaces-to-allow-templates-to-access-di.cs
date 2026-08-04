using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Custom helper class placed in a separate namespace.
    namespace Helpers
    {
        public static class MyHelper
        {
            public static string GetMessage()
            {
                return "Hello from MyHelper!";
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists.
            const string outputFolder = "Output";
            Directory.CreateDirectory(outputFolder);

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a tag that accesses a static field from System.Math.
            builder.Writeln("PI value: <<[Math.PI]>>");

            // Insert a tag that calls a static method from the custom helper class.
            builder.Writeln("Custom message: <<[Helpers.MyHelper.GetMessage()]>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputFolder, "Template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and configure the ReportingEngine.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Register external types so that the template can reference them.
            engine.KnownTypes.Add(typeof(System.Math));
            engine.KnownTypes.Add(typeof(Helpers.MyHelper));

            // No data source is required for this example; an empty object is sufficient.
            object dummyRoot = new object();

            // Build the report. The root name is irrelevant because the template does not reference it.
            engine.BuildReport(doc, dummyRoot, "root");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(outputFolder, "Result.docx");
            doc.Save(resultPath);
        }
    }
}
