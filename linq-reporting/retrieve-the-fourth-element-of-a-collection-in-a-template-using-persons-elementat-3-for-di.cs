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

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
    public ReportModel()
    {
        // Sample data with at least four persons
        Persons.Add(new Person("Alice", 30));
        Persons.Add(new Person("Bob", 25));
        Persons.Add(new Person("Charlie", 28));
        Persons.Add(new Person("Diana", 32));
        Persons.Add(new Person("Ethan", 27));
    }
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and insert the LINQ Reporting tag that accesses the fourth element.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The tag uses ElementAt(3) to get the fourth person (zero‑based index).
        builder.Writeln("Fourth person: <<[model.Persons.ElementAt(3).Name]>> (Age: <<[model.Persons.ElementAt(3).Age]>>)");

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
