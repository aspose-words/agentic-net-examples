using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert placeholders that reference the root object named "person".
        builder.Writeln("First name: <<[person.FirstName]>>");
        builder.Writeln("Last name : <<[person.LastName]>>");
        builder.Writeln("Age       : <<[person.Age]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template back (simulating a real‑world scenario where the template is stored separately).
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare the data source – a single Person instance.
        Person person = new Person
        {
            FirstName = "John",
            LastName = "Doe",
            Age = 30
        };

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "person", therefore we pass it as the third argument.
        engine.BuildReport(loadedTemplate, person, "person");

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}

// Simple data model that matches the placeholders used in the template.
public class Person
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName  { get; set; } = string.Empty;
    public int    Age       { get; set; }
}
