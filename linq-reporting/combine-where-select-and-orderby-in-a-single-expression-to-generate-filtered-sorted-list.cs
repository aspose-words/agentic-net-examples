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

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered and Sorted Persons:");
        builder.Writeln("<<foreach [person in Persons.Where(p => p.Age > 30).Select(p => p).OrderBy(p => p.Name)]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 40 },
                new Person { Name = "David", Age = 25 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
