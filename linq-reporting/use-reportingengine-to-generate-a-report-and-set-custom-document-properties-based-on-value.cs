using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Wrapper model that provides additional values for the template.
    public class ReportModel
    {
        // The date/time when the report is generated.
        public DateTime GeneratedOn { get; set; } = DateTime.Now;
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- 1. Prepare sample data in a DataSet ----------
            DataSet dataSet = new DataSet("ReportData");
            DataTable employeesTable = new DataTable("Employees");
            employeesTable.Columns.Add("Name", typeof(string));
            employeesTable.Columns.Add("Title", typeof(string));
            employeesTable.Columns.Add("Department", typeof(string));

            employeesTable.Rows.Add("Alice Johnson", "Senior Engineer", "R&D");
            employeesTable.Rows.Add("Bob Smith", "Project Manager", "Operations");
            employeesTable.Rows.Add("Carol White", "Analyst", "Finance");

            dataSet.Tables.Add(employeesTable);

            // ---------- 2. Create a template document with LINQ Reporting tags ----------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Employee Report");
            // Use the wrapper model to output the generation date.
            builder.Writeln("Generated on: <<[model.GeneratedOn]>>");
            builder.Writeln(); // empty line

            // Iterate over the Employees table from the DataSet.
            // Because the DataSet will be passed with the name "data",
            // the tag must reference the table via that name.
            builder.Writeln("<<foreach [emp in data.Employees]>>");
            builder.Writeln("Name: <<[emp.Name]>>");
            builder.Writeln("Title: <<[emp.Title]>>");
            builder.Writeln("Department: <<[emp.Department]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required by the lifecycle rule).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // ---------- 3. Load the template and build the report ----------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };

            // Prepare the wrapper model instance.
            ReportModel model = new ReportModel();

            // Build the report using both the model and the DataSet.
            // The first data source is named "model", the second is named "data".
            engine.BuildReport(report,
                new object[] { model, dataSet },
                new string[] { "model", "data" });

            // ---------- 4. Set custom document properties based on the DataSet ----------
            int employeeCount = employeesTable.Rows.Count;
            report.CustomDocumentProperties.Add("EmployeeCount", employeeCount);

            if (employeeCount > 0)
            {
                string firstEmployeeName = employeesTable.Rows[0]["Name"]?.ToString() ?? string.Empty;
                report.CustomDocumentProperties.Add("FirstEmployeeName", firstEmployeeName);
            }

            // ---------- 5. Save the final report ----------
            const string reportPath = "Report.docx";
            report.Save(reportPath);
        }
    }
}
