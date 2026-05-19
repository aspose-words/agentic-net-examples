using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the final report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // 1. Create the LINQ Reporting template.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        // This line may become empty if Age is null, producing an empty paragraph.
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 3. Prepare sample data with some missing values.
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = null },   // Missing age -> empty paragraph.
                new Person { Name = "Charlie", Age = 25 }
            }
        };

        // 4. Configure the ReportingEngine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // 5. Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the final document.
        reportDoc.Save(reportPath);
    }
}

// Wrapper class that matches the root name used in the template.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data entity.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int? Age { get; set; }
}
