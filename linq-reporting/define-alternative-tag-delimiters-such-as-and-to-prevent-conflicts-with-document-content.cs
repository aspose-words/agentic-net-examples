using System;
using System.Collections.Generic;
using System.Text;
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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document using the default LINQ Reporting delimiters.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report using default tag delimiters");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine (no special options needed for default delimiters).
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
