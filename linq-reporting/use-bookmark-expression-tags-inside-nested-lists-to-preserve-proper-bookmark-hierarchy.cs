using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    public class Category
    {
        public string Name { get; set; } = "";
        public string Bookmark { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string Bookmark { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Fruits",
                        Bookmark = "bmFruits",
                        Items = new List<Item>
                        {
                            new Item { Name = "Apple",  Bookmark = "bmApple" },
                            new Item { Name = "Banana", Bookmark = "bmBanana" }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Bookmark = "bmVegetables",
                        Items = new List<Item>
                        {
                            new Item { Name = "Carrot", Bookmark = "bmCarrot" },
                            new Item { Name = "Tomato", Bookmark = "bmTomato" }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // Step 1: Create the template document programmatically
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Report with nested lists and bookmarks:");
            builder.Writeln("<<foreach [category in Categories]>>");
            // Category bookmark and name
            builder.Writeln("<<bookmark [category.Bookmark]>><<[category.Name]>> <</bookmark>>");
            // Items under the category
            builder.Writeln("<<foreach [item in category.Items]>>");
            builder.Writeln("\t- <<bookmark [item.Bookmark]>><<[item.Name]>> <</bookmark>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model"
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
