using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportingEngineExample
{
    static void Main()
    {
        // Load a template document that contains tags referencing a data source named "Employees".
        // Example tag in the template: <<[Employees.Name]>>
        Document template = new Document("Template.docx");

        // Prepare a DataSet with two tables: Employees and Departments.
        DataSet dataSet = new DataSet();

        // Employees table
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));
        employees.Rows.Add("John Doe", 30);
        employees.Rows.Add("Jane Smith", 28);
        dataSet.Tables.Add(employees);

        // Departments table (not used in this example, but shows multiple collections)
        DataTable departments = new DataTable("Departments");
        departments.Columns.Add("DeptName", typeof(string));
        departments.Rows.Add("HR");
        departments.Rows.Add("IT");
        dataSet.Tables.Add(departments);

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the overload that accepts a collection name.
        // The third argument ("Employees") tells the engine which table in the DataSet
        // should be referenced by the template tags.
        bool success = engine.BuildReport(template, dataSet, "Employees");

        // Optionally check if parsing succeeded (useful when InlineErrorMessages option is set).
        if (!success)
        {
            Console.WriteLine("Report building failed due to template syntax errors.");
        }

        // Save the populated document.
        template.Save("ReportResult.docx");
    }
}
