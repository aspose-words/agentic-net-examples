using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Root data model containing a collection to bind.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class used in the collection.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the DOCX template with LINQ Reporting tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Sample LINQ Reporting");
            // Opening tag for the collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Content that will be repeated for each item.
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            // Closing tag for the collection.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Prepare sample data to bind to the template.
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Apple" },
                    new Item { Index = 2, Name = "Banana" },
                    new Item { Index = 3, Name = "Cherry" }
                }
            };

            // -------------------------------------------------
            // 3. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Bind the model to the template using the root name "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
