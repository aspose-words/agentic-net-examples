using System;
using System.Collections.Generic;
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
        // 1. Create a template document using the default tag delimiters << and >>.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a foreach loop that iterates over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template back (required before building the report).
        Document doc = new Document(templatePath);

        // 3. Prepare sample data.
        List<Person> persons = new()
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob",   Age = 25 },
            new Person { Name = "Carol", Age = 28 }
        };

        // 4. Configure the ReportingEngine (no custom delimiters needed).
        ReportingEngine engine = new ReportingEngine();

        // 5. Build the report. The root object name must match the tag reference ("Persons").
        engine.BuildReport(doc, persons, "Persons");

        // 6. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
