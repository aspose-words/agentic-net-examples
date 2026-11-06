using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public List<string> Items { get; set; } = new();
    public string SecondToLastItem { get; set; } = string.Empty;

    public ReportModel()
    {
        // Sample collection.
        Items = new List<string> { "Apple", "Banana", "Cherry", "Date" };

        // Fetch the second‑to‑last item using ElementAt and the collection length.
        if (Items.Count >= 2)
        {
            SecondToLastItem = Items.ElementAt(Items.Count - 2);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Second to last item: <<[model.SecondToLastItem]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "output.docx");
        doc.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }
}
