using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Prepare sample data.
        // -----------------------------------------------------------------
        var person = new Person
        {
            Name = "John Doe",
            Age = 42
        };

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Normal data fields.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");

        // Attempt to access a member of a restricted type (System.Type).
        // This expression will be blocked because System.Type is restricted.
        builder.Writeln("Type: <<[person.GetType().FullName]>>");

        // Save the template.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template (required before building the report).
        // -----------------------------------------------------------------
        var loadedDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine security.
        // -----------------------------------------------------------------
        // Restrict access to System.Type and its members.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Create the engine and set the required option.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        // The root object name used in the template is "person".
        engine.BuildReport(loadedDoc, person, "person");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Report.docx");
        loadedDoc.Save(resultPath);

        Console.WriteLine($"Report generated at: {resultPath}");
    }
}

// ---------------------------------------------------------------------
// Simple data model used by the template.
// ---------------------------------------------------------------------
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
