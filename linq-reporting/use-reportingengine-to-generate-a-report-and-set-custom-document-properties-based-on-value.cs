using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -------------------- Prepare sample data --------------------
        DataSet dataSet = new DataSet();

        DataTable employeesTable = new DataTable("Employees");
        employeesTable.Columns.Add("Name", typeof(string));
        employeesTable.Columns.Add("Position", typeof(string));
        employeesTable.Columns.Add("Salary", typeof(decimal));

        employeesTable.Rows.Add("John Doe", "Manager", 75000m);
        employeesTable.Rows.Add("Jane Smith", "Developer", 65000m);
        employeesTable.Rows.Add("Bob Johnson", "Tester", 55000m);

        dataSet.Tables.Add(employeesTable);

        // -------------------- Create a template document --------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Employee Report");
        // Insert the current date directly; using a LINQ Reporting tag for DateTime causes a parsing error.
        builder.Writeln($"Generated on: {DateTime.Now}");
        builder.Writeln();

        // Loop through the Employees table using LINQ Reporting tags.
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Position: <<[emp.Position]>>");
        builder.Writeln("Salary: <<[emp.Salary]>>");
        builder.Writeln("<</foreach>>");

        // -------------------- Build the report --------------------
        ReportingEngine engine = new ReportingEngine();
        // The data source name ("ds") can be used in the template if needed.
        engine.BuildReport(template, dataSet, "ds");

        // -------------------- Set custom document properties based on the data --------------------
        int employeeCount = employeesTable.Rows.Count;
        decimal totalSalary = 0m;
        foreach (DataRow row in employeesTable.Rows)
        {
            totalSalary += (decimal)row["Salary"];
        }

        // Add custom properties to the generated document.
        // Use the overload that accepts a double for the decimal value.
        template.CustomDocumentProperties.Add("EmployeeCount", employeeCount);
        template.CustomDocumentProperties.Add("TotalSalary", (double)totalSalary);

        // -------------------- Save the final report --------------------
        template.Save("EmployeeReport.docx");
    }
}
