using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<string> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that uses ElementAt to fetch the third item (index 2).
        builder.Writeln("Third item: <<[model.Items.ElementAt(2)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel();
        model.Items.AddRange(new[] { "Apple", "Banana", "Cherry", "Date" });

        // Step 3: Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.Options = ReportBuildOptions.None;

        // The root object name in the template is "model".
        engine.BuildReport(doc, model, "model");

        // Step 4: Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
