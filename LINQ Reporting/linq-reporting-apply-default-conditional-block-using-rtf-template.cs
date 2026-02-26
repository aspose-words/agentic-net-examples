using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingRtfExample
{
    class Program
    {
        static void Main()
        {
            // Path to the RTF template that contains a default conditional block,
            // e.g. <<if [data.ShowText]>>Visible Text<<else>>Hidden Text<<endif>>
            string templatePath = @"C:\Templates\ConditionalBlockTemplate.rtf";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare a simple data source. The property name must match the one used in the template.
            var dataSource = new
            {
                ShowText = true,               // Change to false to test the else branch.
                Text = "Hello, Aspose.Words!"
            };

            // Create and configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after the conditional block is evaluated.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The second parameter is the data source object,
            // the third parameter is the name used to reference the source inside the template.
            engine.BuildReport(template, dataSource, "data");

            // Save the populated document as RTF using custom save options if needed.
            RtfSaveOptions saveOptions = new RtfSaveOptions
            {
                // Example: keep the document generator name in the output.
                ExportGeneratorName = true,
                // Example: pretty‑format the RTF for readability.
                PrettyFormat = true
            };

            string outputPath = @"C:\Output\ConditionalBlockResult.rtf";
            template.Save(outputPath, saveOptions);
        }
    }
}
