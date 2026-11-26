using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataSet.
        DataSet dataSet = new DataSet();
        DataTable salesTable = new DataTable("Sales");
        salesTable.Columns.Add("Month", typeof(string));
        salesTable.Columns.Add("Amount", typeof(int));
        salesTable.Rows.Add("January", 120);
        salesTable.Rows.Add("February", 80);
        salesTable.Rows.Add("March", 150);
        salesTable.Rows.Add("April", 60);
        dataSet.Tables.Add(salesTable);

        // Create the template document.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sales Chart");
        builder.Writeln("<<foreach [row in data.Sales]>>");

        // Start the table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Month");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.InsertCell();
        builder.Writeln("Bar");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[row.Month]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Amount]>>");
        builder.InsertCell();
        // HTML bar whose width is proportional to the amount.
        builder.Writeln(
            "<<html [\"<div style='background:Blue;height:10px;width:' + (row.Amount * 5) + 'px'></div>\"]>>");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build the report using the DataSet as the data source.
        engine.BuildReport(reportDoc, dataSet, "data");

        // Save the final report.
        string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
