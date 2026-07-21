using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample CSV file.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllText(csvPath,
            "Id,Name,Status\r\n" +
            "1,Alice,Open\r\n" +
            "2,Bob,Closed\r\n" +
            "3,Charlie,Open\r\n" +
            "4,David,InProgress\r\n" +
            "5,Eve,Closed\r\n");

        // 2. Build a template document containing LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Status Summary Report");
        // Use GroupBy in the template to aggregate rows by the Status column.
        builder.Writeln("<<foreach [g in persons.GroupBy(r => r[\"Status\"].ToString())]>>");
        builder.Writeln("Status: <<[g.Key]>> - Count: <<[g.Count()]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Create a CSV data source with header row detection.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "persons");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workDir, "ReportOutput.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine("Report generated at: " + outputPath);
    }
}
