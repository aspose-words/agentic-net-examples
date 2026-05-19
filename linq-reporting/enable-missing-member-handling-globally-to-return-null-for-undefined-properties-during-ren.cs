using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Defined property.
    public string Name { get; set; } = "John Doe";

    // No Age property – this will be missing.
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document and insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The template references a missing member (Age). With AllowMissingMembers it will be treated as null.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>"); // Age does not exist on Person.

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Enable global handling of missing members – they will be treated as null.
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Optional: customize the message shown for a plain missing member reference.
        // Leaving it empty results in an empty string in the output.
        engine.MissingMemberMessage = string.Empty;

        // Build the report using the document, the data source, and the root name "person".
        Person data = new Person();
        engine.BuildReport(doc, data, "person");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "MissingMemberReport.docx");
        doc.Save(outputPath);
    }
}
