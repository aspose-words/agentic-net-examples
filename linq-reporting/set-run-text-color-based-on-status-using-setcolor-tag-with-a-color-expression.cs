using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class RunItem
{
    public string Status { get; set; } = "";
    public string Description { get; set; } = "";
}

public class ReportModel
{
    public List<RunItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Task Status Report");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Use the textColor tag. The color expression selects a color based on the item's status.
        // Green for Completed, Orange for Pending, Red for any other status.
        builder.Writeln(
            "<<textColor [item.Status == \"Completed\" ? \"Green\" : item.Status == \"Pending\" ? \"Orange\" : \"Red\"]>>" +
            "Status: <<[item.Status]>> - <<[item.Description]>> " +
            "<</textColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data.
        var model = new ReportModel
        {
            Items = new List<RunItem>
            {
                new RunItem { Status = "Completed", Description = "Generate invoice" },
                new RunItem { Status = "Pending",   Description = "Review contract" },
                new RunItem { Status = "Failed",   Description = "Deploy update" }
            }
        };

        // Create the reporting engine and generate the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        doc.Save(reportPath);
    }
}
