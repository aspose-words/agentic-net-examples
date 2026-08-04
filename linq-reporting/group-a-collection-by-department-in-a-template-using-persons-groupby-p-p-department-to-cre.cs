using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}

// Wrapper class required by the ReportingEngine (cannot be anonymous).
public class ReportData
{
    // Property name must match the name used inside the template tags.
    public List<Person> persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var persons = new List<Person>
        {
            new() { Name = "Alice",   Age = 30, Department = "HR" },
            new() { Name = "Bob",     Age = 45, Department = "Finance" },
            new() { Name = "Charlie", Age = 28, Department = "HR" },
            new() { Name = "Diana",   Age = 35, Department = "IT" },
            new() { Name = "Evan",    Age = 40, Department = "Finance" }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Outer foreach – groups by department.
        builder.Writeln("<<foreach [deptGroup in persons.GroupBy(p => p.Department)]>>");
        builder.Writeln("Department: <<[deptGroup.Key]>>");

        // Inner foreach – iterates over persons in the current group.
        builder.Writeln("<<foreach [p in deptGroup]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Prepare the data source wrapper.
        var data = new ReportData { persons = persons };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, data);

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
