using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBackgroundColor
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the final report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple title.
            builder.Writeln("Product List");

            // Begin a foreach loop over the collection Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Use the backColor tag. The expression [item.Color] evaluates to a color name or code.
            // The content inside the tag (item name and price) will be rendered with that background.
            builder.Writeln("<<backColor [item.Color]>>Item: <<[item.Name]>> - Price: $<<[item.Price]>> <</backColor>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for reporting.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data model.
            // -------------------------------------------------
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apple",  Price = 1.20, Color = "LightYellow" },
                    new Item { Name = "Banana", Price = 0.80, Color = "LightGoldenrodYellow" },
                    new Item { Name = "Cherry", Price = 2.50, Color = "LightCoral" }
                }
            };

            // -------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting Engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }

    // Root data model referenced in the template as <<[model...]>>
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Individual item displayed in the foreach loop.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
        public string Color { get; set; } = string.Empty; // e.g., "LightYellow"
    }
}
