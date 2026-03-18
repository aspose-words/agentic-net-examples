using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

namespace HyperlinkValidationExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains hyperlink fields.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");

            // Load the template document.
            Document doc = new Document(templatePath);

            // Find all hyperlink fields in the document.
            var hyperlinkFields = doc.Range.Fields
                .OfType<FieldHyperlink>()
                .ToList();

            // Validate each hyperlink field for an empty target.
            for (int i = 0; i < hyperlinkFields.Count; i++)
            {
                FieldHyperlink hyperlink = hyperlinkFields[i];

                // A hyperlink may have its address in either the Address or Target property.
                // Consider it empty if both are null, empty, or whitespace.
                bool isEmptyTarget = string.IsNullOrWhiteSpace(hyperlink.Address) &&
                                     string.IsNullOrWhiteSpace(hyperlink.Target);

                if (isEmptyTarget)
                {
                    // Log the error with sufficient context.
                    Console.WriteLine($"Error: Hyperlink field at index {i} has an empty target.");
                }
            }

            // Prepare a simple data source for the reporting engine.
            // Replace this with your actual data source as needed.
            var dataSource = new { };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportOutput.docx");
            doc.Save(outputPath);
        }
    }
}
