using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the markdown template that contains a table with a conditional column.
            // Example of a markdown table in the template (Template.md):
            // | Name | <<if [ds.ShowValue]>><<[ds.Value]>>><<endif>> |
            // The conditional block will render the "Value" column only when ShowValue is true.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.md");

            // Load the markdown template into an Aspose.Words Document.
            Document doc = new Document(templatePath);

            // Prepare the data source. The root object contains a collection named "Items".
            // Each item has a Name, a Value, and a boolean flag ShowValue that controls the column.
            var dataSource = new
            {
                Items = new[]
                {
                    new { Name = "Apple",  Value = 5,  ShowValue = true  },
                    new { Name = "Banana", Value = 3,  ShowValue = false },
                    new { Name = "Cherry", Value = 7,  ShowValue = true  }
                }
            };

            // Create the ReportingEngine and configure options.
            // RemoveEmptyParagraphs ensures that rows/columns that become empty are removed.
            // InlineErrorMessages helps debugging template syntax issues.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs |
                          ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The data source name "ds" is used inside the template.
            // The template can reference the collection with <<foreach [ds.Items]>> and fields with <<[ds.Name]>> etc.
            engine.BuildReport(doc, dataSource, "ds");

            // Save the generated document to DOCX format.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
