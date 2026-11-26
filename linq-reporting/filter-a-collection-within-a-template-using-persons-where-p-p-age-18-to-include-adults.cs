using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Sample data with mixed ages.
        List<Person> persons = new()
        {
            new Person { Name = "Alice", Age = 25 },
            new Person { Name = "Bob", Age = 17 },
            new Person { Name = "Charlie", Age = 30 }
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Adults (Age > 18):");
        builder.Writeln("<<foreach [p in persons.Where(p => p.Age > 18)]>>");
        builder.Writeln("- <<[p.Name]>> (<<[p.Age]>> years)");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, demonstrates the workflow).
        template.Save("Template.docx");

        // Load the template for report generation.
        Document report = new Document("Template.docx");

        // Build the report using the persons list as the root data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, persons, "persons");

        // Save the generated report.
        report.Save("Report.docx");
    }
}
