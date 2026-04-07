using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(workDir, "Data");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with European date format (dd/MM/yyyy).
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(dataDir, "sample.csv");
        string[] csvLines =
        {
            "Date,Value",
            "31/12/2022,123",
            "01/01/2023,456"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 2. Create a Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(dataDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use a foreach block to iterate over the CSV rows.
        builder.Writeln("<<foreach [record in records]>>");
        builder.Writeln("Date (parsed): <<[record.Date]>>");
        builder.Writeln("Value: <<[record.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template back (required by the lifecycle rule).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Set a custom culture (e.g., French) before loading the CSV.
        //    This influences type parsing for dates in the CSV data source.
        // -----------------------------------------------------------------
        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR"); // European format dd/MM/yyyy

        // -----------------------------------------------------------------
        // 5. Configure CSV load options (headers are present).
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true
        };

        // -----------------------------------------------------------------
        // 6. Create the CSV data source using the custom culture.
        // -----------------------------------------------------------------
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 7. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "records");

        // -----------------------------------------------------------------
        // 8. Restore the original thread culture.
        // -----------------------------------------------------------------
        CultureInfo.CurrentCulture = originalCulture;

        // -----------------------------------------------------------------
        // 9. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }
}
