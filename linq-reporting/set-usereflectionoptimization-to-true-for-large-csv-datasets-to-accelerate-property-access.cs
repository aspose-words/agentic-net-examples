using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Id,Name,Value\r\n1,Apple,10\r\n2,Banana,20\r\n3,Cherry,30", Encoding.UTF8);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("CSV Data Report");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Value: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Enable reflection optimization for large data sets.
        ReportingEngine.UseReflectionOptimization = true;

        // Load CSV data source.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, csvData, "data");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
