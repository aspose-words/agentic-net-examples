using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Empty value – the paragraph containing only this tag will be removed.
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words in some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a blank document and a builder to insert the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Paragraph that contains only a LINQ Reporting tag.
            builder.Writeln("<<[model.Name]>>");

            // Paragraph with regular text – it should remain after the report is built.
            builder.Writeln("Static paragraph that stays.");

            // Prepare the data source. The Name property is empty, so the tag resolves to an empty string.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to remove empty paragraphs after tag processing.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            doc.Save("ReportWithRemovedEmptyParagraphs.docx");
        }
    }
}
