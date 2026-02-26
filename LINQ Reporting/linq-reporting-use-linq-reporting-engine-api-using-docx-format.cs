using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting Engine template.
        // The template iterates over a collection named "person" and prints each person's Name and Age.
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person]>>");
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Age: <<[Age]>>");
        builder.Writeln("--------------------");
        builder.Writeln("<</foreach>>");

        // Prepare the data source – a list of Person objects.
        List<Person> persons = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 }
        };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, persons, "person");

        // Save the resulting document in DOCX format.
        doc.Save("PeopleReport.docx");
    }

    // Simple POCO class that serves as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
