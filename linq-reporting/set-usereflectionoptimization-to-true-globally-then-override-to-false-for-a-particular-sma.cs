using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = string.Empty;
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output files.
        string templatePath = "Template.docx";
        string largeReportPath = "ReportLarge.docx";
        string smallReportPath = "ReportSmall.docx";

        // Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document template = new Document(templatePath);

        // Prepare a large data source.
        Model largeModel = new Model();
        for (int i = 1; i <= 100; i++)
        {
            largeModel.Persons.Add(new Person { Name = $"Person {i}" });
        }

        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Build report for the large data source.
        ReportingEngine engineLarge = new ReportingEngine();
        engineLarge.BuildReport(template, largeModel, "model");
        template.Save(largeReportPath);

        // Prepare a small data source.
        Model smallModel = new Model
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice" },
                new Person { Name = "Bob" }
            }
        };

        // Disable reflection optimization for the small data source.
        ReportingEngine.UseReflectionOptimization = false;

        // Reload the template to avoid residual state.
        Document templateForSmall = new Document(templatePath);

        // Build report for the small data source.
        ReportingEngine engineSmall = new ReportingEngine();
        engineSmall.BuildReport(templateForSmall, smallModel, "model");
        templateForSmall.Save(smallReportPath);
    }
}
