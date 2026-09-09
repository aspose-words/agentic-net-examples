using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data.
        List<Person> persons = new()
        {
            new Person { Name = "Alice", Age = 30, Department = "HR" },
            new Person { Name = "Bob", Age = 45, Department = "IT" },
            new Person { Name = "Charlie", Age = 28, Department = "HR" },
            new Person { Name = "Diana", Age = 35, Department = "Finance" },
            new Person { Name = "Evan", Age = 40, Department = "IT" }
        };

        // Create a template document.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Report grouped by Department:");
        // Outer loop – groups by Department.
        builder.Writeln("<<foreach [deptGroup in persons.GroupBy(p => p.Department)]>>");
        builder.Writeln("Department: <<[deptGroup.Key]>>");
        // Inner loop – persons inside the current group.
        builder.Writeln("<<foreach [p in deptGroup]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        // Pass the collection as a data source named "persons".
        engine.BuildReport(doc, persons, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Simple data model.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}
