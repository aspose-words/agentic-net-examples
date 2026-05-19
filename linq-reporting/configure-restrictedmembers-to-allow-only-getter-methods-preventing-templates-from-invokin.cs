using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Public getter only – no setter is exposed to the template.
    public string Name { get; }
    public int Age { get; }

    public Person(string name, int age)
    {
        Name = name;
        Age = age;
    }
}

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age:  <<[model.Age]>>");
        templateDoc.Save(templatePath);

        // 2. Load the template (demonstrates the load step required by the workflow).
        Document doc = new Document(templatePath);

        // 3. Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Optional: allow missing members to avoid exceptions if the template references a non‑existent member.
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // 4. Prepare the data source – a concrete class (no anonymous types) with only getter properties.
        Person person = new Person("John Doe", 42);

        // 5. Build the report. The root object name used in the template is "model".
        engine.BuildReport(doc, person, "model");

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
