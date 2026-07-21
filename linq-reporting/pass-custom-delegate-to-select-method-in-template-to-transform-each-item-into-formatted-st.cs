using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Collection of raw items.
    public List<string> Items { get; set; } = new();

    public ReportModel()
    {
        // Sample data.
        Items.Add("Apple");
        Items.Add("Banana");
        Items.Add("Cherry");
    }

    // Method that formats a single item – adds a prefix and makes the text uppercase.
    public string FormatItem(string s) => $"Item: {s.ToUpper()}";
}

public class Program
{
    public static void Main()
    {
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("<<[model.FormatItem(item)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportModel model = new ReportModel();

        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
