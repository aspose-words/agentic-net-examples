using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV reading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(workDir, "output");
        Directory.CreateDirectory(outputDir);

        // Create sample CSV file.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllText(csvPath, "Name,Value\r\nAlice,10\r\n,\r\nBob,20");

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Static content that should stay unchanged.
        builder.Writeln("=== Report Start ===");

        // CSV processing section.
        builder.Writeln("<<foreach [row in CsvRows]>>");
        builder.Writeln("<<[row.Name]>> - <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Another static paragraph that must remain even if empty.
        builder.Writeln("=== Report End ===");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Set up CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, new CsvDataLoadOptions { HasHeaders = true });

        // Configure the reporting engine to remove empty paragraphs after processing CSV.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root name "CsvRows" matches the foreach tag.
        engine.BuildReport(doc, csvData, "CsvRows");

        // Save the final document.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }
}
