using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Duration string in the format "hh:mm:ss"
    public string Duration { get; set; } = "02:15:30";
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that parses the duration string into a TimeSpan
        // and then outputs the total minutes.
        // The engine needs to know about the TimeSpan type to call its static Parse method.
        builder.Writeln("Total minutes: <<[TimeSpan.Parse(model.Duration).TotalMinutes]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register TimeSpan so that its static members can be used in the template.
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. The root object name must match the tag prefix ("model").
        engine.BuildReport(doc, model, "model");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the generated document.
        string outputPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outputPath);
    }
}
