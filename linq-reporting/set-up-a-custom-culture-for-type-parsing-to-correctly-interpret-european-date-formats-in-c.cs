using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Set a European culture (French) for the current thread.
        // This culture will be used when parsing dates from the CSV file.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

        // Prepare sample CSV data with a European date format (dd/MM/yyyy).
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Name,Date\r\n" +
            "John,31/12/2020\r\n" +
            "Anna,15/01/2021\r\n");

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tags: iterate over the CSV rows (named Records).
        builder.Writeln("<<foreach [record in Records]>>");
        builder.Writeln("Name: <<[record.Name]>>");
        builder.Writeln("Date: <<[record.Date]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options (comma delimiter, headers present).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '"',
            HasHeaders = true
        };

        // Create the CSV data source using the custom culture set above.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report. The data source name used in the template is "Records".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "Records");

        // Save the generated report.
        reportDoc.Save("report.docx");
    }
}
