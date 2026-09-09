using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Non‑nullable property – initialized to avoid warnings.
        public string Name { get; set; } = "World";

        // This property will be null, causing the corresponding tag to produce an empty value.
        public string? EmptyTag { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document and add a template with LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Paragraph that will always contain text.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Paragraph that contains only a tag. When the tag evaluates to null/empty,
            // the paragraph becomes empty and should be removed.
            builder.Writeln("<<[model.EmptyTag]>>");

            // 2. Configure the ReportingEngine to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // 3. Prepare the data source.
            ReportModel model = new ReportModel(); // EmptyTag remains null.

            // 4. Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(doc, model, "model");

            // 5. Save the resulting document.
            doc.Save("ReportOutput.docx");

            Console.WriteLine("Report generated successfully.");
        }
    }
}
