using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string csvPath = Path.Combine(workDir, "data.csv");
        string outputDir = Path.Combine(workDir, "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple LINQ Reporting template.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Create a small CSV file with headers.
        File.WriteAllText(csvPath, "Name,Age\r\nAlice,30\r\nBob,25\r\nCharlie,35");

        // ---------- Report with reflection optimization enabled ----------
        ReportingEngine.UseReflectionOptimization = true;

        var docWithOpt = new Document(templatePath);
        var csvLoadOptions = new CsvDataLoadOptions { HasHeaders = true };
        var csvDataSource = new CsvDataSource(csvPath, csvLoadOptions);
        var engineOpt = new ReportingEngine();
        engineOpt.BuildReport(docWithOpt, csvDataSource, "data");
        string optOutputPath = Path.Combine(outputDir, "report_optimized.docx");
        docWithOpt.Save(optOutputPath);

        // ---------- Report with reflection optimization disabled ----------
        ReportingEngine.UseReflectionOptimization = false;

        var docWithoutOpt = new Document(templatePath);
        var csvDataSourceNoOpt = new CsvDataSource(csvPath, csvLoadOptions);
        var engineNoOpt = new ReportingEngine();
        engineNoOpt.BuildReport(docWithoutOpt, csvDataSourceNoOpt, "data");
        string noOptOutputPath = Path.Combine(outputDir, "report_no_optimization.docx");
        docWithoutOpt.Save(noOptOutputPath);
    }
}
