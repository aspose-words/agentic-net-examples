using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach block that iterates over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Output a link for each item.
        // If LinkText is not null or empty, use it as the display text.
        // Otherwise output a link with only the URL – the engine will use the URL as the display text.
        builder.Writeln(
            "Link: " +
            "<<if [item.LinkText != null && item.LinkText != \"\"]>>" +
                "<<link [item.Url] [item.LinkText]>>" +
            "<</if>>" +
            "<<if [item.LinkText == null || item.LinkText == \"\"]>>" +
                "<<link [item.Url]>>" +
            "<</if>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        Model data = new Model
        {
            Items = new List<Item>
            {
                new Item
                {
                    Url = "https://www.aspose.com",
                    LinkText = "Aspose Home"
                },
                new Item
                {
                    Url = "https://www.github.com",
                    LinkText = null // No display text; URL will be used.
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Allow missing members so that a null LinkText does not cause an error.
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.BuildReport(template, data, "model");

        // Save the generated document.
        template.Save("Report.docx");
    }
}

// Root data model referenced in the template as "model".
public class Model
{
    public List<Item> Items { get; set; } = new();
}

// Individual item containing a URL and an optional display text.
public class Item
{
    public string Url { get; set; } = string.Empty;
    public string? LinkText { get; set; }
}
