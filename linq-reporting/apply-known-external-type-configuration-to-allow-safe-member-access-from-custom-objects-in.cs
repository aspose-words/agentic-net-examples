using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple data field.
        builder.Writeln("Name: <<[model.Name]>>");

        // Use a known external static type (Helper) to format the date.
        builder.Writeln("Birth Date: <<[Helper.FormatDate(model.BirthDate)]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data source.
        Person person = new Person
        {
            Name = "John Doe",
            BirthDate = new DateTime(1990, 5, 23)
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the external type so its static members can be used safely in the template.
        engine.KnownTypes.Add(typeof(Helper));

        // Build the report. The root object name must match the name used in the template tags ("model").
        engine.BuildReport(loadedTemplate, person, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the template.
// ---------------------------------------------------------------------
public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime BirthDate { get; set; }
}

// ---------------------------------------------------------------------
// External static helper class whose members are allowed in the template.
// ---------------------------------------------------------------------
public static class Helper
{
    // Formats a DateTime as a short date string.
    public static string FormatDate(DateTime date) => date.ToString("yyyy-MM-dd");
}
