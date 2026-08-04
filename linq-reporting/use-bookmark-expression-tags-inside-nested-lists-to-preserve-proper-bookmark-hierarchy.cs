using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Use a bullet list for the hierarchy.
        builder.ListFormat.List = template.Lists.Add(ListTemplate.BulletDefault);

        // Top‑level foreach over categories.
        builder.Writeln("<<foreach [category in Model.Categories]>>");
        builder.Writeln("<<bookmark [category.Bookmark]>>Category: <<[category.Name]>> <</bookmark>>");

        // Indent for items.
        builder.ListFormat.ListIndent();
        builder.Writeln("<<foreach [item in category.Items]>>");
        builder.Writeln("<<bookmark [item.Bookmark]>>Item: <<[item.Name]>> <</bookmark>>");

        // Indent for sub‑items.
        builder.ListFormat.ListIndent();
        builder.Writeln("<<foreach [sub in item.SubItems]>>");
        builder.Writeln("<<bookmark [sub.Bookmark]>>SubItem: <<[sub.Name]>> <</bookmark>>");
        builder.Writeln("<</foreach>>"); // end sub‑items foreach
        builder.ListFormat.ListOutdent(); // outdent sub‑items

        builder.Writeln("<</foreach>>"); // end items foreach
        builder.ListFormat.ListOutdent(); // outdent items

        builder.Writeln("<</foreach>>"); // end categories foreach
        builder.ListFormat.RemoveNumbers(); // stop list formatting

        // Save the template (optional, shown for clarity).
        const string templatePath = "BookmarkTemplate.docx";
        template.Save(templatePath);

        // Load the template (demonstrates the load step required before BuildReport).
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Bookmark = "Bookmark_Fruits",
                    Items = new List<Item>
                    {
                        new Item
                        {
                            Name = "Apple",
                            Bookmark = "Bookmark_Apple",
                            SubItems = new List<SubItem>
                            {
                                new SubItem { Name = "Red Apple", Bookmark = "Bookmark_RedApple" },
                                new SubItem { Name = "Green Apple", Bookmark = "Bookmark_GreenApple" }
                            }
                        },
                        new Item
                        {
                            Name = "Banana",
                            Bookmark = "Bookmark_Banana",
                            SubItems = new List<SubItem>
                            {
                                new SubItem { Name = "Ripe Banana", Bookmark = "Bookmark_RipeBanana" }
                            }
                        }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Bookmark = "Bookmark_Vegetables",
                    Items = new List<Item>
                    {
                        new Item
                        {
                            Name = "Carrot",
                            Bookmark = "Bookmark_Carrot",
                            SubItems = new List<SubItem>()
                        }
                    }
                }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "Model");

        // Save the generated report.
        const string outputPath = "BookmarkReport.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
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
    public List<SubItem> SubItems { get; set; } = new();
}

public class SubItem
{
    public string Name { get; set; } = "";
    public string Bookmark { get; set; } = "";
}
