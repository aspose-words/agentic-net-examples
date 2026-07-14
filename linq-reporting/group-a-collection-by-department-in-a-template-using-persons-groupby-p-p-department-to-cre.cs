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
    public string Department { get; set; } = "";
}

public class ReportModel
{
    public List<Person> persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // 2. Prepare sample data.
        ReportModel model = new ReportModel
        {
            persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, Department = "HR" },
                new Person { Name = "Bob", Age = 45, Department = "Finance" },
                new Person { Name = "Charlie", Age = 28, Department = "HR" },
                new Person { Name = "Diana", Age = 35, Department = "IT" },
                new Person { Name = "Ethan", Age = 40, Department = "Finance" }
            }
        };

        // 3. Load the template and build the report.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 4. Save the generated report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Employee Report");
        builder.Writeln();

        // Group by Department
        builder.Writeln("<<foreach [deptGroup in persons.GroupBy(p => p.Department)]>>");
        builder.Writeln("Department: <<[deptGroup.Key]>>");
        builder.Writeln();

        // List persons within the current department
        builder.Writeln("<<foreach [p in deptGroup]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // End outer foreach
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}
