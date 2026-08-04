using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that filter the collection to adults (Age > 18).
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > 18)]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for report generation.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var data = new Model
        {
            Persons = new List<Person>
            {
                new() { Name = "Alice", Age = 25 },
                new() { Name = "Bob", Age = 17 },
                new() { Name = "Charlie", Age = 30 },
                new() { Name = "Diana", Age = 15 }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, data, "model");

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
