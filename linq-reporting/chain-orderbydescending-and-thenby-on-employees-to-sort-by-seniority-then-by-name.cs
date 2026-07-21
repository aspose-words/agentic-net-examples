using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Seniority { get; set; }
}

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Employees Report:");
        // The foreach tag iterates over the Employees collection sorted by Seniority (desc) then Name (asc).
        builder.Writeln("<<foreach [emp in Employees.OrderByDescending(e => e.Seniority).ThenBy(e => e.Name)]>>");
        builder.Writeln("Name: <<[emp.Name]>>, Seniority: <<[emp.Seniority]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template (simulating a separate load step).
        var loadedTemplate = new Document(templatePath);

        // 3. Prepare sample data.
        var employees = new List<Employee>
        {
            new Employee { Name = "Alice",   Seniority = 5 },
            new Employee { Name = "Bob",     Seniority = 7 },
            new Employee { Name = "Charlie", Seniority = 5 },
            new Employee { Name = "David",   Seniority = 9 }
        };

        var model = new ReportModel { Employees = employees };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // 5. Save the generated report.
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
