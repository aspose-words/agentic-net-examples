using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Original data model (contains the Salary property that we want to hide).
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        public decimal Salary { get; set; } = 75000m; // This member will be hidden.
    }

    // Wrapper model used for the report – it deliberately omits the Salary property.
    public class PersonRestricted
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public PersonRestricted(Person source)
        {
            Name = source.Name;
            Age = source.Age;
            // Salary is not exposed, so the engine will treat it as a missing member.
        }
    }

    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Salary: <<[person.Salary]>>"); // This field will be restricted.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // 2. Load the template (simulating a separate load step).
        Document loadedTemplate = new Document(templatePath);

        // 3. Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Allow missing members so that restricted members are treated as empty.
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "[Hidden]";

        // 4. Build the report using the restricted wrapper model.
        Person data = new Person();
        PersonRestricted restrictedData = new PersonRestricted(data);
        engine.BuildReport(loadedTemplate, restrictedData, "person");

        // 5. Save the generated report.
        string resultPath = Path.Combine(outputDir, "Result.docx");
        loadedTemplate.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }
}
