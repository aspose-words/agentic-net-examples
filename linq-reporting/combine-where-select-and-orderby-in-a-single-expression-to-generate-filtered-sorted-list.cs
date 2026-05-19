using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered and Sorted Names:");
        // Combine Where, OrderBy, and Select in a single expression inside the foreach tag.
        builder.Writeln("<<foreach [name in model.People.Where(p => p.Age > 30).OrderBy(p => p.Name).Select(p => p.Name)]>>");
        builder.Writeln(" - <<[name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data model with sample data.
        // -------------------------------------------------
        ReportModel model = new()
        {
            People = new()
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 35 },
                new Person { Name = "Dave",  Age = 22 },
                new Person { Name = "Eve",   Age = 40 }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using Aspose.Words ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple scenario.
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
