using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Name may be null in the data source.
    public string? Name { get; set; }
}

public class ReportModel
{
    // Collection of persons to be iterated in the template.
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory("Output");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a foreach loop that iterates over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        // The Name field may be null; the engine will treat it as an empty string.
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine("Output", "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data with a missing collection item (null Name).
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = null },          // Missing value.
                new Person { Name = "Alice" },
                new Person { Name = "Bob" }
            }
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine.
        //    The default behavior already treats null members as empty strings,
        //    so no special flag is required.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options needed.

        // Build the report. The root object name must match the template references.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine("Output", "Report.docx");
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {outputPath}");
    }
}
