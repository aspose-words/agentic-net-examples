using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportItem
{
    // URL of the hyperlink.
    public string Url { get; set; } = "";
    // Display text for the hyperlink.
    public string Text { get; set; } = "";
}

public class ReportModel
{
    // Collection that will be iterated in the template.
    public List<ReportItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel();
        model.Items.Add(new ReportItem { Url = "https://www.example.com", Text = "Example Site" });
        model.Items.Add(new ReportItem { Url = "https://www.github.com", Text = "GitHub" });

        // -----------------------------------------------------------------
        // Step 1: Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table inside the foreach block.
        var table = builder.StartTable();

        // First cell: place a link tag that will become a functional hyperlink.
        builder.InsertCell();
        builder.Writeln("<<link [item.Url] [item.Text]>>");

        // End the single row and the table.
        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report using the model; the root name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
