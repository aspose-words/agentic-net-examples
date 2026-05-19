using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a DOCX template with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("Date: <<[model.GeneratedDate]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("  <<[item.Index]>> - <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("HTML content:");
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        builder.Writeln("<<textColor [\"DarkBlue\"]>>Report Completed<</textColor>>");

        // Save the template to disk (required before loading for report generation).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Title = "Sample LINQ Reporting",
            GeneratedDate = DateTime.Now.ToString("yyyy-MM-dd"),
            HtmlSnippet = "<p style=\"color:red;\">This is <b>HTML</b> snippet.</p>",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the populated document as PDF.
        doc.Save("Report.pdf", SaveFormat.Pdf);
    }
}

// Public data model classes.
public class ReportModel
{
    public string Title { get; set; } = "";
    public string GeneratedDate { get; set; } = "";
    public string HtmlSnippet { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
