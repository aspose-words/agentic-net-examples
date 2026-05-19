using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    // Simple employee data model.
    public class Employee
    {
        public string Name { get; set; } = "";
        public string Title { get; set; } = "";
        public string Department { get; set; } = "";
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // LINQ Reporting tags:
            // Group employees by Department using GroupBy, then list each employee.
            builder.Writeln("<<foreach [dept in model.Employees.GroupBy(e => e.Department)]>>");
            builder.Writeln("Department: <<[dept.Key]>>");
            builder.Writeln("<<foreach [emp in dept]>>");
            builder.Writeln("- <<[emp.Name]>> (<<[emp.Title]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee { Name = "Alice Johnson", Title = "Developer", Department = "IT" },
                    new Employee { Name = "Bob Smith", Title = "Analyst", Department = "Finance" },
                    new Employee { Name = "Carol White", Title = "Developer", Department = "IT" },
                    new Employee { Name = "David Brown", Title = "Manager", Department = "HR" },
                    new Employee { Name = "Eve Davis", Title = "Analyst", Department = "Finance" }
                }
            };

            // 3. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
