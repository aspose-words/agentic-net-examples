using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBookmarkExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    public class Category
    {
        public string Name { get; set; } = "";
        public string BookmarkName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // -------------------------------------------------
            // Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Report with Nested Bookmarks");
            builder.Writeln();

            // Begin outer foreach for categories
            builder.Writeln("<<foreach [cat in Model.Categories]>>");

            // Apply numbered list for categories
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("<<bookmark [cat.BookmarkName]>><<[cat.Name]>> <</bookmark>>");

            // Indent for inner list (items)
            builder.ListFormat.ListIndent();

            // Begin inner foreach for items
            builder.Writeln("<<foreach [item in cat.Items]>>");
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("<<bookmark [item.BookmarkName]>><<[item.Name]>> <</bookmark>>");
            builder.Writeln("<</foreach>>");

            // Outdent back to outer level
            builder.ListFormat.ListOutdent();

            // End outer foreach
            builder.Writeln("<</foreach>>");

            // Remove any list formatting
            builder.ListFormat.RemoveNumbers();

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // Prepare sample data
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Fruits",
                        BookmarkName = "Bookmark_Fruits",
                        Items = new List<Item>
                        {
                            new Item { Name = "Apple", BookmarkName = "Bookmark_Apple" },
                            new Item { Name = "Banana", BookmarkName = "Bookmark_Banana" }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        BookmarkName = "Bookmark_Vegetables",
                        Items = new List<Item>
                        {
                            new Item { Name = "Carrot", BookmarkName = "Bookmark_Carrot" },
                            new Item { Name = "Lettuce", BookmarkName = "Bookmark_Lettuce" }
                        }
                    }
                }
            };

            // -------------------------------------------------
            // Build the report using LINQ Reporting Engine
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "Model");

            // Save the generated report
            reportDoc.Save(outputPath);
        }
    }
}
