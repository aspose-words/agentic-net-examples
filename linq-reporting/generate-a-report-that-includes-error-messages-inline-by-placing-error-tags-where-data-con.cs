using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = "";
    public string AgeString { get; set; } = "";
    public string DateString { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data with values that will cause conversion errors.
        var model = new ReportModel
        {
            Name = "John Doe",
            AgeString = "NotANumber",          // int.Parse will fail
            DateString = "InvalidDateValue"    // DateTime.Parse will fail
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[int.Parse(model.AgeString)]>>");
        builder.Writeln("Date: <<[DateTime.Parse(model.DateString)]>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report. The method returns a flag indicating whether parsing succeeded.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportWithInlineErrors.docx";
        doc.Save(outputPath);

        // Optionally, write the result to the console (no interactive input required).
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}
