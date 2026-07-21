using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for the reporting engine.
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare sample hierarchical data.
        var model = new ReportModel
        {
            Companies = new List<Company>
            {
                new Company
                {
                    Name = "Tech Corp",
                    Departments = new List<Department>
                    {
                        new Department
                        {
                            Name = "Research",
                            Employees = new List<Employee>
                            {
                                new Employee { Name = "Alice Johnson", JoinDate = new DateTime(2018, 3, 12) },
                                new Employee { Name = "Bob Smith", JoinDate = new DateTime(2019, 7, 23) }
                            }
                        },
                        new Department
                        {
                            Name = "Development",
                            Employees = new List<Employee>
                            {
                                new Employee { Name = "Carol White", JoinDate = new DateTime(2020, 1, 5) }
                            }
                        }
                    }
                },
                new Company
                {
                    Name = "Innovate Ltd",
                    Departments = new List<Department>
                    {
                        new Department
                        {
                            Name = "Marketing",
                            Employees = new List<Employee>
                            {
                                new Employee { Name = "David Brown", JoinDate = new DateTime(2017, 11, 30) }
                            }
                        }
                    }
                }
            }
        };

        // Create the template document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Root object name is "model".
        builder.Writeln("<<foreach [c in model.Companies]>>");
        builder.Writeln("Company: <<[c.Name]>>");
        builder.Writeln("<<foreach [d in c.Departments]>>");
        builder.Writeln("  Department: <<[d.Name]>>");
        builder.Writeln("  <<foreach [e in d.Employees]>>");
        builder.Writeln("    Employee: <<[StringHelper.Upper(e.Name)]>> - Joined: <<[DateHelper.FormatDate(e.JoinDate)]>>");
        builder.Writeln("  <</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // Register external types used in the template.
        engine.KnownTypes.Add(typeof(DateHelper));
        engine.KnownTypes.Add(typeof(StringHelper));

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Company> Companies { get; set; } = new();
}

// Company entity.
public class Company
{
    public string Name { get; set; } = string.Empty;
    public List<Department> Departments { get; set; } = new();
}

// Department entity.
public class Department
{
    public string Name { get; set; } = string.Empty;
    public List<Employee> Employees { get; set; } = new();
}

// Employee entity.
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public DateTime JoinDate { get; set; }
}

// Helper class for date formatting.
public static class DateHelper
{
    public static string FormatDate(DateTime date) => date.ToString("yyyy-MM-dd");
}

// Helper class for string manipulation.
public static class StringHelper
{
    public static string Upper(string value) => value?.ToUpperInvariant() ?? string.Empty;
}
