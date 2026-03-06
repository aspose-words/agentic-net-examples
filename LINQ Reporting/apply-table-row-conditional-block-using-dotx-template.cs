using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace TableRowConditionalExample
{
    // Simple data source class used by the template.
    public class RowData
    {
        public bool ShowRow { get; set; }
        public string Description { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template that contains a conditional block inside a table row.
            // The template should have a tag like <<if [row.ShowRow]>> ... <<endif>> surrounding the row.
            Document template = new Document(@"Templates\ConditionalTableRow.dotx");

            // Prepare the data source. The ReportingEngine will evaluate the condition for each row.
            RowData[] rows = new RowData[]
            {
                new RowData { ShowRow = true,  Description = "First visible row" },
                new RowData { ShowRow = false, Description = "This row will be hidden" },
                new RowData { ShowRow = true,  Description = "Second visible row" }
            };

            // Build the report by merging the data source with the template.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("rows") must match the name used in the template tags.
            engine.BuildReport(template, rows, "rows");

            // Save the resulting document.
            template.Save(@"Results\ConditionalTableResult.docx");
        }
    }
}
