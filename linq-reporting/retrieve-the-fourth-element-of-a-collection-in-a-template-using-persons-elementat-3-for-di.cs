using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
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
        // Prepare sample data – a list of at least four persons.
        List<Person> persons = new()
        {
            new Person("Alice", 30),
            new Person("Bob", 25),
            new Person("Charlie", 28),
            new Person("Diana", 32),
            new Person("Ethan", 27)
        };

        // Create a template document programmatically.
        Document doc = new();
        DocumentBuilder builder = new(doc);
        builder.Writeln("Fourth person in the list: <<[persons.ElementAt(3).Name]>>");
        builder.Writeln("Age: <<[persons.ElementAt(3).Age]>>");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, persons, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
