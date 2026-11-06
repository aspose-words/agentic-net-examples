using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Age { get; set; } = "";
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Persons =
            {
                new Person { Id = "1", Name = "Alice", Age = "30" },
                new Person { Id = "2", Name = "Bob",   Age = "25" }
            }
        };

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a foreach loop that iterates over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("ID: <<[p.Id]>>, Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the model as the data source.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
