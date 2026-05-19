using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> People { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("People Report");
        // Use a foreach tag that iterates over the collection directly.
        builder.Writeln("<<foreach [p in model.People]>>");
        // Format each item inside the loop.
        builder.Writeln("- <<[p.Name]>> (<<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // 2. Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            People = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
