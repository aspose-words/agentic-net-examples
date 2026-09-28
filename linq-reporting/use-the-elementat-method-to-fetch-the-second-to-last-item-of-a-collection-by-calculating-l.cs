using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<string> Items { get; set; } = new();

    // Optional helper property that uses ElementAt to get the second‑to‑last item.
    public string? SecondToLast => Items.Count >= 2 ? Items.ElementAt(Items.Count - 2) : null;
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to insert the template content.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple paragraph that displays the second‑to‑last item using ElementAt.
        // The expression calculates the index based on the collection length.
        builder.Writeln("Second to last item (using ElementAt): <<[model.Items.ElementAt(model.Items.Count - 2)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template back for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<string> { "Alpha", "Beta", "Gamma", "Delta", "Epsilon" }
        };

        // Build the report using Aspose.Words LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
