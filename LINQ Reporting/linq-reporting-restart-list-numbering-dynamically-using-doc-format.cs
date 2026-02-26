using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsListRestartExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains a numbered list.
            Document doc = new Document("Template.docx");

            // Assume the first list in the document is the one we want to control.
            // Enable restarting of the list numbering at each section.
            if (doc.Lists.Count > 0)
            {
                List list = doc.Lists[0];
                list.IsRestartAtEachSection = true;
            }

            // Prepare a data source for the reporting engine.
            // This can be any .NET object; here we use an anonymous type for demonstration.
            var dataSource = new
            {
                Items = new[]
                {
                    new { Text = "First item" },
                    new { Text = "Second item" },
                    new { Text = "Third item" }
                }
            };

            // Build the report using the loaded template and the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "ds");

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
