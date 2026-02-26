using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalColumn
{
    // Simple data model used by the LINQ Reporting Engine.
    public class ReportItem
    {
        // Determines whether the column should be displayed.
        public bool ShowColumn { get; set; }

        // Value to be placed in the column when it is shown.
        public string Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template document that contains the DOT syntax.
            // The template should have a table with a column like:
            // <<if [ShowColumn]>><<[Value]>><<endif>>
            string templatePath = @"C:\Templates\ConditionalColumnTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a data source with a mix of rows that show and hide the column.
            List<ReportItem> data = new List<ReportItem>
            {
                new ReportItem { ShowColumn = true,  Value = "First"  },
                new ReportItem { ShowColumn = false, Value = "Second" }, // Column will be omitted.
                new ReportItem { ShowColumn = true,  Value = "Third"  }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional blocks are omitted.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the data source.
            // The engine will evaluate the <<if>> blocks for each row.
            engine.BuildReport(doc, data, "item");

            // Save the populated document.
            string outputPath = @"C:\Output\ConditionalColumnReport.docx";
            doc.Save(outputPath);
        }
    }
}
