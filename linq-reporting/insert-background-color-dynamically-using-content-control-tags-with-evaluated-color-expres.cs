using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

namespace LinqReportingBackgroundColorExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public bool IsHighlighted { get; set; }

        // Evaluated color expression for background color
        public string BackgroundColor => IsHighlighted ? "LightGray" : "White";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Create a template document programmatically
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Items List:");

            // Begin foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table with one column to display item name with dynamic background color
            Table table = builder.StartTable();
            builder.InsertCell();
            // Apply backColor tag with evaluated color expression
            builder.Writeln("<<backColor [item.BackgroundColor]>> <<[item.Name]>> <</backColor>>");
            builder.EndRow();
            builder.EndTable();

            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template for reporting
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Sample data
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Name = "Alpha",   IsHighlighted = true },
                    new Item { Name = "Beta",    IsHighlighted = false },
                    new Item { Name = "Gamma",   IsHighlighted = true },
                    new Item { Name = "Delta",   IsHighlighted = false }
                }
            };

            // Build the report using LINQ Reporting engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine(outputDir, "output.docx");
            doc.Save(outputPath);
        }
    }
}
