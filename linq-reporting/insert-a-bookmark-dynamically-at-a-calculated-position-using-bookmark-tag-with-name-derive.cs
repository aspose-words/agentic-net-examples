using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkLinqReportingExample
{
    // Data model classes
    public class ReportModelWrapper
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Title { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            const string templatePath = "template.docx";
            const string outputPath = "output.docx";

            // -------------------------------------------------
            // Create the template document with LINQ Reporting tags
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");
            // Insert a bookmark whose name comes from the data field
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            // The content of the bookmark – the title of the item
            builder.Writeln("<<[item.Title]>>");
            // Close the bookmark and the foreach block
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Prepare sample data
            // -------------------------------------------------
            var model = new ReportModelWrapper
            {
                Items = new List<Item>
                {
                    new Item { Title = "First Chapter", BookmarkName = "BM_First" },
                    new Item { Title = "Second Chapter", BookmarkName = "BM_Second" },
                    new Item { Title = "Conclusion", BookmarkName = "BM_Conclusion" }
                }
            };

            // -------------------------------------------------
            // Load the template and build the report
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model; the root name is "model"
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated document
            reportDoc.Save(outputPath);
        }
    }
}
