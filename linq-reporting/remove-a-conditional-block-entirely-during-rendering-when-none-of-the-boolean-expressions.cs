using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalRemoval
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // When true the conditional block will be rendered, otherwise it will be removed.
        public bool ShowMessage { get; set; } = false; // Initialized to avoid nullable warnings.
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write a paragraph that contains a conditional block.
            // If ShowMessage is false the block evaluates to false and the paragraph becomes empty.
            // With the RemoveEmptyParagraphs option the empty paragraph will be removed entirely.
            builder.Writeln("<<if [model.ShowMessage]>>This message is shown only when ShowMessage is true.<</if>>");

            // Save the template to a local file (optional, shown for clarity).
            const string templatePath = "ConditionalTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template (demonstrates the load step required by the rules).
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            ReportModel model = new ReportModel
            {
                ShowMessage = false // Change to true to see the block rendered.
            };

            // 4. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // RemoveEmptyParagraphs ensures that paragraphs that become empty after tag processing are deleted.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // 5. Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(doc, model, "model");

            // 6. Save the resulting document.
            const string outputPath = "ConditionalReport.docx";
            doc.Save(outputPath);

            // The program finishes without waiting for user input.
        }
    }
}
