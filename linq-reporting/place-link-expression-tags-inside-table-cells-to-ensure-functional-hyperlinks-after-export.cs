using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Links = new List<LinkInfo>
            {
                new LinkInfo { Url = "https://www.example.com",   LinkText = "Example Site" },
                new LinkInfo { Url = "https://www.github.com",    LinkText = "GitHub" },
                new LinkInfo { Url = "https://www.microsoft.com", LinkText = "Microsoft" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin the foreach block that iterates over the Links collection.
        // Use the root name "model" when referencing the collection.
        builder.Writeln("<<foreach [link in model.Links]>>");

        // Create a table that will be repeated for each item.
        Table table = builder.StartTable();

        // Insert a cell for the hyperlink.
        builder.InsertCell();
        builder.Writeln("<<link [link.Url] [link.LinkText]>>");

        // End the row for this iteration.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of link items to be displayed in the table.
    public List<LinkInfo> Links { get; set; } = new();
}

public class LinkInfo
{
    // URL the hyperlink points to.
    public string Url { get; set; } = string.Empty;

    // Text displayed for the hyperlink.
    public string LinkText { get; set; } = string.Empty;
}
