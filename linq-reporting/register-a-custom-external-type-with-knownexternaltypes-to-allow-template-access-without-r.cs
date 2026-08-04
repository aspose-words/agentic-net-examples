using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use a foreach loop over model.Persons and call a static helper method.
        builder.Writeln("<<foreach [p in model.Persons]>>");
        builder.Writeln("<<[MyHelper.GetGreeting(p.Name)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice" },
                new Person { Name = "Bob" },
                new Person { Name = "Charlie" }
            }
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Register the external type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(MyHelper));

        // Build the report using the root object name "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine("Report generated successfully.");
    }
}

// ---------------------------------------------------------------------
// Helper class whose static members are accessed from the template.
// ---------------------------------------------------------------------
public static class MyHelper
{
    // Returns a greeting message for the supplied name.
    public static string GetGreeting(string name) => $"Hello, {name}!";
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Person
{
    // Name of the person.
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    // Collection of persons to iterate over in the template.
    public List<Person> Persons { get; set; } = new();
}
