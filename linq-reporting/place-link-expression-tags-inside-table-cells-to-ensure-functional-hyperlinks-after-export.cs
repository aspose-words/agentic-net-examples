using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinkItem
{
    public string Url { get; set; } = "";
    public string Text { get; set; } = "";
}

public class ReportModel
{
    public List<LinkItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new LinkItem { Url = "https://www.example.com", Text = "Example Site" },
                new LinkItem { Url = "https://www.github.com", Text = "GitHub" },
                new LinkItem { Url = "https://www.microsoft.com", Text = "Microsoft" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach block that iterates over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table where each row contains a hyperlink.
        var table = builder.StartTable();

        // First (and only) cell of the row.
        builder.InsertCell();
        // Insert the link tag. The first expression is the URL, the second is the display text.
        builder.Writeln("<<link [item.Url] [item.Text]>>");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report. Hyperlinks will be functional.
        // -----------------------------------------------------------------
        var outputDocx = "Report.docx";
        var outputPdf = "Report.pdf";

        reportDoc.Save(outputDocx);
        reportDoc.Save(outputPdf);
    }
}
