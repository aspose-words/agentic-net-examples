using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedLoops
{
    // Data model classes
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    public class Category
    {
        public string Name { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for template and output documents
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Step 1: Create the template document with nested foreach tags
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Outer foreach over categories
            builder.Writeln("<<foreach [cat in Categories]>>");
            builder.Writeln("Category: <<[cat.Name]>>");
            builder.Writeln("");

            // Inner foreach over items within a category
            builder.Writeln("<<foreach [item in cat.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Description]>>");
            builder.Writeln("<</foreach>>");

            // Close outer foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Prepare sample hierarchical data
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Fruits",
                        Items = new List<Item>
                        {
                            new Item { Name = "Apple", Description = "Red or green fruit" },
                            new Item { Name = "Banana", Description = "Long yellow fruit" }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Items = new List<Item>
                        {
                            new Item { Name = "Carrot", Description = "Orange root vegetable" },
                            new Item { Name = "Lettuce", Description = "Leafy green" }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // Step 3: Load the template and build the report
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model as the root object named "model"
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            reportDoc.Save(outputPath);
        }
    }
}
