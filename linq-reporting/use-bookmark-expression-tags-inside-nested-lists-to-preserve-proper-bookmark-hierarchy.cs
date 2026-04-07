using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Root data model for the report.
    public class ReportModel
    {
        // Collection of sections; each section will have its own bookmark.
        public List<Section> Sections { get; set; } = new();
    }

    // Represents a section that contains a title and a list of items.
    public class Section
    {
        public string Title { get; set; } = string.Empty;
        // Name of the bookmark that will be created for the section title.
        public string TitleBookmark { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    // Represents an item inside a section; each item also gets a bookmark.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        // Name of the bookmark that will be created for the item.
        public string BookmarkName { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Optional heading.
            builder.Writeln("Report generated with Aspose.Words LINQ Reporting");
            builder.Writeln();

            // Begin looping over sections.
            builder.Writeln("<<foreach [section in Sections]>>");

            // Section title wrapped in a bookmark.
            builder.Writeln("<<bookmark [section.TitleBookmark]>>");
            builder.Writeln("<<[section.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln();

            // Begin looping over items inside the current section.
            builder.Writeln("<<foreach [item in section.Items]>>");

            // Each item is wrapped in its own bookmark.
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            builder.Writeln("• <<[item.Name]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln();

            // End of items foreach.
            builder.Writeln("<</foreach>>");

            // End of sections foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data that matches the template structure.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "Section A",
                        TitleBookmark = "SecA",
                        Items = new List<Item>
                        {
                            new Item { Name = "Item 1", BookmarkName = "SecA_Item1" },
                            new Item { Name = "Item 2", BookmarkName = "SecA_Item2" }
                        }
                    },
                    new Section
                    {
                        Title = "Section B",
                        TitleBookmark = "SecB",
                        Items = new List<Item>
                        {
                            new Item { Name = "Item 3", BookmarkName = "SecB_Item3" },
                            new Item { Name = "Item 4", BookmarkName = "SecB_Item4" }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "model".
            engine.BuildReport(report, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            report.Save(outputPath);

            // The program finishes without waiting for user input.
        }
    }
}
