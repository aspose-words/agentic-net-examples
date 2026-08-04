using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        // Non‑nullable properties must be initialized to avoid warnings.
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;

        // This property returns null, causing the corresponding paragraph to become empty.
        public string? Empty { get; set; } = null;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Result.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            // This line will become empty after processing because the tag resolves to null.
            builder.Writeln("<<[person.Empty]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine to remove empty paragraphs.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Root data object.
            Person model = new Person();

            // Build the report. The root name in the template is "person".
            engine.BuildReport(doc, model, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(resultPath);
        }
    }
}
