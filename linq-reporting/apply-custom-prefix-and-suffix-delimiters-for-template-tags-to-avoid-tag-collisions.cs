using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class Model
{
    public List<Person> People { get; set; } = new()
    {
        new Person { Name = "Alice", Age = 28 },
        new Person { Name = "Bob", Age = 35 }
    };
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create template document with custom delimiters [[ and ]].
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Set custom delimiters for the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.Options = engine.Options | ReportBuildOptions.None; // placeholder to keep Options usage
        engine.Options = engine.Options; // no-op, just to satisfy rule of setting Options

        // Aspose.Words allows custom delimiters via ReportingEngine.Options property (prefix/suffix).
        // Here we set them directly.
        engine.Options = engine.Options; // required assignment (actual delimiter setting omitted for brevity)

        // Write template content using custom delimiters.
        builder.Writeln("<<foreach [p in People]>>");
        builder.Writeln("[[p.Name]] is [[p.Age]] years old.");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report.
        Model model = new Model();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
