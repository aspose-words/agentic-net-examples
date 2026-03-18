using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace NestedForeachExample
{
    // Simple data model: a Category has a Name and a collection of Items.
    public class Category
    {
        public string Name { get; set; }
        public List<Item> Items { get; set; } = new List<Item>();
    }

    // Simple data model: an Item has a Name.
    public class Item
    {
        public string Name { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Insert a template that uses Aspose.Words Reporting Engine syntax.
            //    The outer foreach iterates over $categories$, the inner foreach iterates over $items$ of each category.
            //    <<[Name]>> refers to the current object's Name property.
            builder.Writeln("Categories and their items:");
            builder.Writeln("<<foreach [in categories]>>");
            builder.Writeln("- Category: <<[Name]>>");
            builder.Writeln("  Items:");
            builder.Writeln("  <<foreach [in items]>>");
            builder.Writeln("    * <<[Name]>>");
            builder.Writeln("  <</foreach>>");
            builder.Writeln("<</foreach>>");

            // 3. Prepare hierarchical data.
            var categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple" },
                        new Item { Name = "Banana" },
                        new Item { Name = "Cherry" }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new List<Item>
                    {
                        new Item { Name = "Carrot" },
                        new Item { Name = "Lettuce" },
                        new Item { Name = "Tomato" }
                    }
                }
            };

            // 4. Build the report using the ReportingEngine.
            //    The data source is an anonymous object that exposes the "categories" collection.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new { categories });

            // 5. Save the resulting document.
            doc.Save("NestedForeachReport.docx");
        }
    }
}
