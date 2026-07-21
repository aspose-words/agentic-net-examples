using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
    public decimal Salary { get; set; } = 75000m; // This member will be hidden
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        Person model = new Person();

        // Create a template document with LINQ Reporting tags
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Configure the reporting engine
        ReportingEngine engine = new ReportingEngine();

        // Restrict the Person type so its members cannot be accessed.
        // This will make the Salary tag unavailable while allowing other members.
        ReportingEngine.SetRestrictedTypes(typeof(Person));

        // Allow missing members so that the hidden tag does not cause an exception.
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report using the model object; root name is "person"
        engine.BuildReport(doc, model, "person");

        // Save the generated report
        doc.Save("Report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert tags that reference the model's members
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Salary: <<[person.Salary]>>"); // This will be hidden by the restriction

        doc.Save(filePath);
    }
}
