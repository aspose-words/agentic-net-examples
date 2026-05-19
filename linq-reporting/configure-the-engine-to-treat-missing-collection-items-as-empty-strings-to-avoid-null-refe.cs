using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model.
    public class ReportModel
    {
        // Collection to be iterated in the template.
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class.
    public class Item
    {
        // Name of the item – initialized via constructor.
        public string Name { get; set; }

        public Item(string name) => Name = name;
    }

    public class Program
    {
        public static void Main()
        {
            // Required for some Aspose.Words features.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();

            builder.Writeln("Item list:");
            // foreach loop over Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Use an if‑condition to safely handle null items.
            builder.Writeln("- <<if [item != null]>><<[item.Name]>> <</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            builder.Document.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data with a missing (null) collection item.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item("Apple"),
                    null,                     // Missing item – will be skipped by the if‑condition.
                    new Item("Banana")
                }
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine to treat missing members as empty.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = string.Empty
            };

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            var outputPath = "Report.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
