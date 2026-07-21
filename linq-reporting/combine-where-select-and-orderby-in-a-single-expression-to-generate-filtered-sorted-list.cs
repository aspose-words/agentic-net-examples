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
    // Holds the filtered and sorted list of names.
    public List<string> Names { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        List<Person> people = new()
        {
            new Person { Name = "Alice", Age = 28 },
            new Person { Name = "Bob",   Age = 35 },
            new Person { Name = "Carol", Age = 42 },
            new Person { Name = "Dave",  Age = 31 },
            new Person { Name = "Eve",   Age = 25 }
        };

        // LINQ: filter Age > 30, select Name, order alphabetically.
        List<string> filteredNames = people
            .Where(p => p.Age > 30)
            .Select(p => p.Name)
            .OrderBy(name => name)
            .ToList();

        // Prepare the model for the reporting engine.
        ReportModel model = new() { Names = filteredNames };

        // Create a template document with a foreach tag.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered and sorted names:");
        builder.Writeln("<<foreach [n in Names]>>");
        builder.Writeln(" - <<[n]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the model.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save("Report.docx");
    }
}
