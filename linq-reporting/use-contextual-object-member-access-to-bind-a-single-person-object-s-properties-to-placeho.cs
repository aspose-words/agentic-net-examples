using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public int Age { get; set; }

    public Person(string firstName, string lastName, int age)
    {
        FirstName = firstName;
        LastName = lastName;
        Age = age;
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags that reference the Person object's members.
        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[person.FirstName]>> <<[person.LastName]>>");
        builder.Writeln("Age: <<[person.Age]>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and bind a single Person instance.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Sample data source.
        Person person = new Person("John", "Doe", 30);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, person, "person");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}
