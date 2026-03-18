using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportingWithCustomProperties
{
    static void Main()
    {
        // Load the template document. Replace with the actual path to your .docx template.
        Document template = new Document("Template.docx");

        // Prepare a DataSet that will be used as the data source for the report.
        DataSet dataSet = new DataSet();

        // Example table "Employees" with some columns.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("FirstName", typeof(string));
        employees.Columns.Add("LastName", typeof(string));
        employees.Columns.Add("HireDate", typeof(DateTime));
        employees.Columns.Add("Salary", typeof(double));
        employees.Columns.Add("IsFullTime", typeof(bool));

        // Add sample rows.
        employees.Rows.Add("John", "Doe", new DateTime(2015, 3, 12), 75000.0, true);
        employees.Rows.Add("Jane", "Smith", new DateTime(2018, 7, 1), 62000.0, false);
        dataSet.Tables.Add(employees);

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report by merging the template with the DataSet.
        // The second parameter is the data source; the third (optional) name can be omitted.
        engine.BuildReport(template, dataSet);

        // Set custom document properties based on values from the DataSet.
        // For demonstration, we take the first row values.
        if (employees.Rows.Count > 0)
        {
            DataRow firstRow = employees.Rows[0];

            // String property.
            template.CustomDocumentProperties.Add("ReportAuthor", "ReportingEngine Demo");

            // DateTime property.
            template.CustomDocumentProperties.Add("ReportDate", DateTime.Now);

            // Number (int) property – total number of employees.
            template.CustomDocumentProperties.Add("EmployeeCount", employees.Rows.Count);

            // Double property – average salary.
            double avgSalary = 0;
            foreach (DataRow row in employees.Rows)
                avgSalary += (double)row["Salary"];
            avgSalary /= employees.Rows.Count;
            template.CustomDocumentProperties.Add("AverageSalary", avgSalary);

            // Boolean property – whether the first employee is full‑time.
            template.CustomDocumentProperties.Add("FirstEmployeeFullTime", (bool)firstRow["IsFullTime"]);
        }

        // Save the generated report.
        template.Save("ReportWithProperties.docx");
    }
}
