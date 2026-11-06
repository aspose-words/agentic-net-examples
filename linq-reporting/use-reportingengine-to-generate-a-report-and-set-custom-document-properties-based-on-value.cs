using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    // Title of the report.
    public string ReportTitle { get; set; } = "Employee Report";

    // Collection of employee rows.
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data in a DataSet.
        // -----------------------------------------------------------------
        DataSet dataSet = new();
        DataTable employeeTable = new("Employees");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Age", typeof(int));
        employeeTable.Rows.Add("Alice", 30);
        employeeTable.Rows.Add("Bob", 45);
        employeeTable.Rows.Add("Charlie", 28);
        dataSet.Tables.Add(employeeTable);

        // -----------------------------------------------------------------
        // 2. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Title placeholder.
        builder.Writeln("<<[model.ReportTitle]>>");
        builder.Writeln();

        // Table header (static part).
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();
        builder.EndTable();

        // Data rows using LINQ Reporting foreach.
        builder.Writeln("<<foreach [emp in model.Employees]>>");
        Table dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[emp.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[emp.Age]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);

        // Prepare the root model.
        ReportModel model = new()
        {
            Employees = employeeTable.AsEnumerable()
                .Select(r => new Employee
                {
                    Name = r.Field<string>("Name") ?? "",
                    Age = r.Field<int>("Age")
                })
                .ToList()
        };

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Set custom document properties based on the data.
        // -----------------------------------------------------------------
        int totalEmployees = model.Employees.Count;
        double averageAge = totalEmployees > 0
            ? model.Employees.Average(e => e.Age)
            : 0;

        reportDoc.CustomDocumentProperties.Add("TotalEmployees", totalEmployees);
        reportDoc.CustomDocumentProperties.Add("AverageAge", averageAge);

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
