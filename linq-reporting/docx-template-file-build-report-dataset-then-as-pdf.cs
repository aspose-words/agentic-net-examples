using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains reporting tags (e.g. <<[DataSet.Table.Column]>>)
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the generated PDF will be saved
            string outputPdfPath = @"C:\Reports\GeneratedReport.pdf";

            // Load the template document from file
            Document templateDoc = new Document(templatePath);

            // Build a sample DataSet with one DataTable and some rows
            DataSet reportData = CreateSampleDataSet();

            // Create a ReportingEngine instance
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with data from the DataSet
            // This will replace the template tags with actual values
            engine.BuildReport(templateDoc, reportData);

            // Save the populated document as PDF
            templateDoc.Save(outputPdfPath, SaveFormat.Pdf);
        }

        // Helper method to create a sample DataSet for the report
        private static DataSet CreateSampleDataSet()
        {
            DataSet ds = new DataSet();

            // Create a DataTable named "Employees"
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("FirstName", typeof(string));
            employees.Columns.Add("LastName", typeof(string));
            employees.Columns.Add("Position", typeof(string));
            employees.Columns.Add("Salary", typeof(decimal));

            // Add sample rows
            employees.Rows.Add("John", "Doe", "Manager", 75000m);
            employees.Rows.Add("Jane", "Smith", "Developer", 62000m);
            employees.Rows.Add("Bob", "Johnson", "Designer", 58000m);

            // Add the table to the DataSet
            ds.Tables.Add(employees);

            return ds;
        }
    }
}
