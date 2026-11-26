using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    // Data model for an employee.
    public class Employee
    {
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
        public string Title { get; set; } = "";
    }

    // Root object passed to the reporting engine.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee { Name = "Alice",   Department = "HR",      Title = "Manager"   },
                    new Employee { Name = "Bob",     Department = "IT",      Title = "Developer" },
                    new Employee { Name = "Charlie", Department = "HR",      Title = "Recruiter" },
                    new Employee { Name = "David",   Department = "IT",      Title = "SysAdmin"  },
                    new Employee { Name = "Eve",     Department = "Finance", Title = "Analyst"   }
                }
            };

            // Create the template document programmatically.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Directory");
            builder.Writeln();

            // Group employees by Department using LINQ GroupBy inside the template.
            builder.Writeln("<<foreach [dept in model.Employees.GroupBy(e => e.Department)]>>");
            builder.Writeln("Department: <<[dept.Key]>>");
            builder.Writeln("<<foreach [emp in dept]>>");
            builder.Writeln("- <<[emp.Name]>> (<<[emp.Title]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
