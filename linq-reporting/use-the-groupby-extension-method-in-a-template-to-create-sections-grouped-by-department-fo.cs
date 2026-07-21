using System;
using System.Collections.Generic;
using System.Linq; // Needed for GroupBy
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

    // Wrapper model that contains the collection used by the report.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Prepare sample data ----------
            var model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Department = "Sales", Title = "Sales Manager" },
                    new Employee { Name = "Jane Smith", Department = "Sales", Title = "Sales Representative" },
                    new Employee { Name = "Bob Johnson", Department = "HR", Title = "HR Specialist" },
                    new Employee { Name = "Alice Brown", Department = "HR", Title = "Recruiter" },
                    new Employee { Name = "Tom Clark", Department = "IT", Title = "Developer" }
                }
            };

            // ---------- Create the LINQ Reporting template ----------
            const string templateFile = "Template.docx";
            var templateDoc = new Document();               // Create a blank document
            var builder = new DocumentBuilder(templateDoc); // Builder for the template

            builder.Writeln("Employee Report");
            // Group employees by Department using GroupBy in the foreach expression.
            builder.Writeln("<<foreach [dept in Employees.GroupBy(e => e.Department)]>>");
            builder.Writeln("Department: <<[dept.Key]>>");
            builder.Writeln("<<foreach [emp in dept]>>");
            builder.Writeln("- <<[emp.Name]>> (<<[emp.Title]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template before building the report.
            templateDoc.Save(templateFile);

            // ---------- Load the template and generate the report ----------
            var reportDoc = new Document(templateFile);
            var engine = new ReportingEngine();

            // Build the report using the model; no root name is needed.
            engine.BuildReport(reportDoc, model);

            // Save the generated report.
            const string outputFile = "EmployeeReport.docx";
            reportDoc.Save(outputFile);
        }
    }
}
