using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string FirstName { get; set; } = "";
    public string LastName { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Sample data
        var person = new Person { FirstName = "John", LastName = "Doe" };

        // Create a template document with an expression tag that concatenates first and last name
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("<<[person.FirstName + \" \" + person.LastName]>>");

        // Build the report using the LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the generated report
        doc.Save("Report.docx");
    }
}
