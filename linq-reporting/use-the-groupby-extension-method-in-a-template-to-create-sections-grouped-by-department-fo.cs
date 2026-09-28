using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Paths for template and final report
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Create the template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Employee Directory");
        builder.Writeln();

        // Begin grouping by Department using GroupBy in the foreach tag
        builder.Writeln("<<foreach [dept in Employees.GroupBy(e => e.Department)]>>");
        builder.Writeln("Department: <<[dept.Key]>>");
        builder.Writeln();

        // List employees within each department
        builder.Writeln("<<foreach [emp in dept]>>");
        builder.Writeln("- <<[emp.Name]>> (<<[emp.Title]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Sample data model
        ReportModel model = new()
        {
            Employees = new()
            {
                new Employee { Name = "Alice Johnson", Title = "Software Engineer", Department = "R&D" },
                new Employee { Name = "Bob Smith", Title = "Senior Engineer", Department = "R&D" },
                new Employee { Name = "Carol White", Title = "HR Manager", Department = "Human Resources" },
                new Employee { Name = "David Brown", Title = "Recruiter", Department = "Human Resources" },
                new Employee { Name = "Eve Davis", Title = "Accountant", Department = "Finance" }
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(reportPath);
    }
}

// Root data model
public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

// Employee data class
public class Employee
{
    public string Name { get; set; } = "";
    public string Title { get; set; } = "";
    public string Department { get; set; } = "";
}
