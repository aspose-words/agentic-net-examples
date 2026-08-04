using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    // Data model for an employee.
    public class Employee
    {
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
    }

    // Wrapper class that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee { Name = "Alice Johnson", Department = "HR" },
                    new Employee { Name = "Bob Smith", Department = "IT" },
                    new Employee { Name = "Carol White", Department = "HR" },
                    new Employee { Name = "David Brown", Department = "Finance" },
                    new Employee { Name = "Eve Davis", Department = "IT" }
                }
            };

            // -----------------------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title.
            builder.Writeln("Employees Report");
            builder.Writeln();

            // GroupBy expression: group employees by Department.
            builder.Writeln("<<foreach [dept in Employees.GroupBy(e => e.Department)]>>");
            builder.Writeln("Department: <<[dept.Key]>>");
            builder.Writeln();

            // List employees within the current department.
            builder.Writeln("<<foreach [emp in dept]>>");
            builder.Writeln("- <<[emp.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // End of outer foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options

            // Build the report using the model; the root name is "model".
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Optionally, you could check the success flag if InlineErrorMessages were enabled.
            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
