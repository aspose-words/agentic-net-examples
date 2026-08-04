using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tag that filters the collection using Where and an external static property.
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > ExternalHelper.MinAge)]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare data source.
        // -------------------------------------------------
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob",   Age = 35 },
                new Person { Name = "Carol", Age = 45 },
                new Person { Name = "Dave",  Age = 28 }
            }
        };

        // -------------------------------------------------
        // 4. Configure ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Register the external type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(ExternalHelper));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);

        // Optional: indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}

// -------------------------------------------------
// Data model classes.
// -------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

// -------------------------------------------------
// External helper class whose static property is used in the LINQ filter.
// -------------------------------------------------
public static class ExternalHelper
{
    // This value can be changed to affect the filtering logic.
    public static int MinAge { get; set; } = 30;
}
