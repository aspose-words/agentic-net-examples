using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string CompanyName { get; set; } = "";
    public DateTime ReportDate { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataSet.
        DataSet dataSet = new DataSet();
        DataTable infoTable = new DataTable("Info");
        infoTable.Columns.Add("CompanyName", typeof(string));
        infoTable.Columns.Add("ReportDate", typeof(DateTime));
        infoTable.Rows.Add("Acme Corp", DateTime.Today);
        dataSet.Tables.Add(infoTable);

        // Create a simple template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report for <<[model.CompanyName]>>");
        builder.Writeln("Date: <<[model.ReportDate]>>");
        // Save the template to disk (required before loading for reporting).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template document.
        Document reportDoc = new Document(templatePath);

        // Map DataSet values to a strongly‑typed model.
        DataRow row = dataSet.Tables["Info"].Rows[0];
        ReportModel model = new ReportModel
        {
            CompanyName = row["CompanyName"].ToString(),
            ReportDate = (DateTime)row["ReportDate"]
        };

        // Generate the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Set custom document properties based on the DataSet values.
        reportDoc.CustomDocumentProperties.Add("CompanyName", model.CompanyName);
        reportDoc.CustomDocumentProperties.Add("ReportDate", model.ReportDate);

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
