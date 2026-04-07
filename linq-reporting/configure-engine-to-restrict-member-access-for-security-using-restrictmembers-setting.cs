using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple foreach loop to list persons.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> is <<[p.Age]>> years old.");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template back (required before building the report).
        Document reportDoc = new Document(templatePath);

        // 3. Prepare sample data.
        Model model = new Model
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // 4. Configure the reporting engine.
        // Restrict access to members of System.Type (as an example of a sensitive type).
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members so that attempts to access restricted members
            // produce empty output instead of throwing an exception.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report using the configured engine.
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
