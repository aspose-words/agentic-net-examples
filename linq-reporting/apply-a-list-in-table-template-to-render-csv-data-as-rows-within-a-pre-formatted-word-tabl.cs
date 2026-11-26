using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        const string csvFileName = "sample.csv";
        File.WriteAllText(csvFileName,
            "Id,Name,Quantity\r\n" +
            "1,Apple,10\r\n" +
            "2,Banana,20\r\n" +
            "3,Cherry,15\r\n");

        // Create the Word template with a pre‑formatted table and LINQ Reporting tags.
        const string templateFileName = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Table header.
        var headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach over CSV rows (root name: Data).
        builder.Writeln("<<foreach [row in Data]>>");

        // Table rows – each iteration creates a new row.
        var dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[row.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templateFileName);

        // Load the template for reporting.
        var doc = new Document(templateFileName);

        // Load CSV data source (first row contains headers).
        var csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        var csvData = new CsvDataSource(csvFileName, csvOptions);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "Data");

        // Save the generated report.
        const string outputFileName = "output.docx";
        doc.Save(outputFileName);
    }
}
