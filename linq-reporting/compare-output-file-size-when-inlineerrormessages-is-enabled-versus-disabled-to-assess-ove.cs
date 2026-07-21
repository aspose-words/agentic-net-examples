using System;
using System.Collections.Generic;
using System.IO;
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
    public ReportModel() { }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();
        model.Persons.Add(new Person("Alice", 30));
        model.Persons.Add(new Person("Bob", 25));
        model.Persons.Add(new Person("Charlie", 35));

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Build report with InlineErrorMessages enabled.
        string reportWithErrorsPath = "Report_WithInlineError.docx";
        BuildReport(templatePath, model, "model", ReportBuildOptions.InlineErrorMessages, reportWithErrorsPath);

        // Build report with InlineErrorMessages disabled (default options).
        string reportWithoutErrorsPath = "Report_WithoutInlineError.docx";
        BuildReport(templatePath, model, "model", ReportBuildOptions.None, reportWithoutErrorsPath);

        // Compare file sizes.
        long sizeWithErrors = new FileInfo(reportWithErrorsPath).Length;
        long sizeWithoutErrors = new FileInfo(reportWithoutErrorsPath).Length;

        Console.WriteLine($"Report size with InlineErrorMessages: {sizeWithErrors} bytes");
        Console.WriteLine($"Report size without InlineErrorMessages: {sizeWithoutErrors} bytes");
        Console.WriteLine($"Size overhead: {sizeWithErrors - sizeWithoutErrors} bytes");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a simple heading.
        builder.Writeln("Person Report");
        builder.Writeln();

        // Insert a foreach loop to list persons.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    private static void BuildReport(string templatePath, ReportModel model, string rootName, ReportBuildOptions options, string outputPath)
    {
        // Load a fresh copy of the template for each build.
        var doc = new Document(templatePath);

        var engine = new ReportingEngine
        {
            Options = options
        };

        // Build the report.
        bool success = engine.BuildReport(doc, model, rootName);
        // success is only meaningful when InlineErrorMessages is set; we ignore it here.

        // Save the generated report.
        doc.Save(outputPath);
    }
}
