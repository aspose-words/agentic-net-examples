using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    public class Employee
    {
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
    }

    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some data sources)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Create the template document with LINQ Reporting tags
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("<<foreach [dept in Employees.GroupBy(e => e.Department)]>>");
            builder.Writeln("Department: <<[dept.Key]>>");
            builder.Writeln("<<foreach [emp in dept]>>");
            builder.Writeln("- <<[emp.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for report generation
            var reportDoc = new Document(templatePath);

            // Sample data
            var model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee { Name = "Alice", Department = "HR" },
                    new Employee { Name = "Bob", Department = "IT" },
                    new Employee { Name = "Charlie", Department = "HR" },
                    new Employee { Name = "David", Department = "Finance" },
                    new Employee { Name = "Eve", Department = "IT" }
                }
            };

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
