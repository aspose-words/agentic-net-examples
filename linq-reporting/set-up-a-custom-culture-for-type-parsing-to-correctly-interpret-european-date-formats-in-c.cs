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
        // Set a custom culture that uses the European date format (dd/MM/yyyy).
        CultureInfo europeanCulture = new CultureInfo("en-GB");
        Thread.CurrentThread.CurrentCulture = europeanCulture;
        Thread.CurrentThread.CurrentUICulture = europeanCulture;

        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample CSV data with European date strings.
        // -----------------------------------------------------------------
        string csvFile = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        string csvContent = "Name,Date\nJohn Doe,31/12/2023\nJane Smith,15/01/2024";
        File.WriteAllText(csvFile, csvContent);

        // -----------------------------------------------------------------
        // 2. Build a simple LINQ Reporting template.
        // -----------------------------------------------------------------
        string templateFile = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Date: <<[row.Date]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templateFile);

        // CSV load options – headers are present, default separator (comma).
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        CsvDataSource csvSource = new CsvDataSource(csvFile, csvOptions);

        // -----------------------------------------------------------------
        // 4. Generate the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvSource, "data");

        // -----------------------------------------------------------------
        // 5. Save the final document.
        // -----------------------------------------------------------------
        string reportFile = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportFile);
    }
}
