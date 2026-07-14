using System;
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

        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(baseDir, "Data");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple CSV file with headers and sample rows.
        string csvPath = Path.Combine(dataDir, "People.csv");
        File.WriteAllText(csvPath,
            "Name,Age\r\n" +
            "Alice,30\r\n" +
            "Bob,25\r\n" +
            "Charlie,35\r\n",
            Encoding.UTF8);

        // 2. Build the template document programmatically.
        DocumentBuilder builder = new DocumentBuilder();
        // The foreach tag iterates over the CSV data source named "csvData".
        builder.Writeln("<<foreach [row in csvData]>>");
        // Inside the loop we can reference column names via the loop variable.
        builder.Writeln("<<[row.Name]>> - <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(dataDir, "Template.docx");
        builder.Document.Save(templatePath);

        // 3. Load the template document.
        Document templateDoc = new Document(templatePath);

        // 4. Create a CsvDataSource that reads the CSV file (headers are present).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // hasHeaders = true
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "csvData".
        engine.BuildReport(templateDoc, csvDataSource, "csvData");

        // 6. Save the generated report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        templateDoc.Save(resultPath);

        // Indicate completion.
        Console.WriteLine("Report generated at: " + resultPath);
    }
}
