using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, Department = "HR" },
                new Person { Name = "Bob",   Age = 25, Department = "IT" },
                new Person { Name = "Carol", Age = 28, Department = "HR" },
                new Person { Name = "Dave",  Age = 35, Department = "Finance" },
                new Person { Name = "Eve",   Age = 22, Department = "IT" }
            }
        };

        // Create a blank document and build the LINQ Reporting template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Employee Report");
        builder.Writeln();

        // Group by Department.
        builder.Writeln("<<foreach [deptGroup in Persons.GroupBy(p => p.Department)]>>");
        builder.Writeln("Department: <<[deptGroup.Key]>>");
        builder.Writeln();

        // List persons within the current department.
        builder.Writeln("<<foreach [p in deptGroup]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // End of outer foreach.
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the root data source named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("EmployeeReport.docx");
    }
}
