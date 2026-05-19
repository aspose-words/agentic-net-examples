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
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -------------------------------------------------
        // 1. Create the template document with LINQ tag.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Adults (Age > 18):");
        // LINQ Reporting tag that filters the collection directly in the template.
        builder.Writeln("<<foreach [p in model.Persons.Where(p => p.Age > 18)]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob", Age = 17 },
                new Person { Name = "Charlie", Age = 30 },
                new Person { Name = "Diana", Age = 15 }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 5. Save the final report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
