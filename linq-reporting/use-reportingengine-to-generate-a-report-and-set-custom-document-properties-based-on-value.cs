using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataSet.
        DataSet dataSet = new DataSet();

        // Employees table.
        DataTable employeesTable = new DataTable("Employees");
        employeesTable.Columns.Add("Name", typeof(string));
        employeesTable.Columns.Add("Age", typeof(int));
        employeesTable.Rows.Add("Alice Johnson", 30);
        employeesTable.Rows.Add("Bob Smith", 45);
        employeesTable.Rows.Add("Carol White", 28);
        dataSet.Tables.Add(employeesTable);

        // CompanyInfo table (used for custom document properties).
        DataTable companyInfoTable = new DataTable("CompanyInfo");
        companyInfoTable.Columns.Add("CompanyName", typeof(string));
        companyInfoTable.Columns.Add("ReportDate", typeof(DateTime));
        companyInfoTable.Rows.Add("Acme Corporation", DateTime.Today);
        dataSet.Tables.Add(companyInfoTable);

        // Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Employee Report");
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>, Age: <<[emp.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSet, "data");

        // Set custom document properties based on values from the DataSet.
        if (dataSet.Tables.Contains("CompanyInfo") && dataSet.Tables["CompanyInfo"].Rows.Count > 0)
        {
            DataRow infoRow = dataSet.Tables["CompanyInfo"].Rows[0];
            string companyName = infoRow["CompanyName"]?.ToString() ?? string.Empty;
            DateTime reportDate = infoRow["ReportDate"] is DateTime dt ? dt : DateTime.MinValue;

            reportDoc.CustomDocumentProperties.Add("CompanyName", companyName);
            reportDoc.CustomDocumentProperties.Add("ReportDate", reportDate);
        }

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
