using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
    public class ReportModel
    {
        // Non‑nullable properties are initialized to avoid warnings.
        public string Name { get; set; } = string.Empty;
        public string? Optional { get; set; } // May be null to produce an empty paragraph.
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[model.Name]>>");
            // This line contains only a tag; if the tag evaluates to null/empty,
            // the paragraph will become empty after rendering.
            builder.Writeln("<<[model.Optional]>>");
            builder.Writeln("End of Report");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Name = "John Doe",
                Optional = null // This will result in an empty paragraph.
            };

            // -----------------------------------------------------------------
            // 3. Build the report with the RemoveEmptyParagraphs option.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(template, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated document.
            // -----------------------------------------------------------------
            const string outputPath = "Report_Output.docx";
            template.Save(outputPath);
        }
    }
}
