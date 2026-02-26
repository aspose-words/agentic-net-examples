using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model for the report.
    public class Item
    {
        public string Name { get; set; }
        public bool Show { get; set; }
    }

    public class ReportData
    {
        public List<Item> Items { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "TableConditionalTemplate.docx";
            const string outputPath = "TableConditionalReport.docx";

            // -----------------------------------------------------------------
            // 1. Create a DOT template that contains a conditional block inside a table row.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Start a table with a single column.
            builder.StartTable();

            // Insert a cell that will be repeated for each item in the collection.
            // The <<foreach>> tag iterates over the Items collection.
            // Inside the row we place an <<if>> block that checks the Show property.
            builder.InsertCell();
            builder.Writeln("<<foreach [Items]>>");               // Begin loop
            builder.Writeln("<<if [Show]>>");                     // Conditional start
            builder.Write("Item: ");
            builder.Writeln("<<[Name]>>");                        // Field to display
            builder.Writeln("<<endif>>");                         // Conditional end
            builder.Writeln("<<endforeach>>");                    // End loop
            builder.EndRow();

            builder.EndTable();

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            var data = new ReportData
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apple",  Show = true  },
                    new Item { Name = "Banana", Show = false },
                    new Item { Name = "Cherry", Show = true  }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using Aspose.Words LINQ ReportingEngine.
            // -----------------------------------------------------------------
            // Load the template.
            Document reportDoc = new Document(templatePath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional blocks are omitted.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source object is passed directly.
            engine.BuildReport(reportDoc, data);

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);

            Console.WriteLine("Report generated successfully.");
        }
    }
}
