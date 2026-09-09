using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public Person(string name, int age)
    {
        Name = name;
        Age = age;
    }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document with a LINQ Reporting tag that accesses the fourth element.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Fourth person: <<[model.Persons.ElementAt(3).Name]>> (Age: <<[model.Persons.ElementAt(3).Age]>>)");
        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data with at least four persons.
        Model data = new Model
        {
            Persons = new List<Person>
            {
                new Person("Alice", 30),
                new Person("Bob", 25),
                new Person("Charlie", 28),
                new Person("Diana", 32),   // Fourth element (index 3)
                new Person("Ethan", 27)
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, data, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
