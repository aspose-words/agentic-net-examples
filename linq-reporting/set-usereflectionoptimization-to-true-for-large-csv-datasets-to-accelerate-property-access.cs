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
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath,
            "Name,Age,City\r\n" +
            "Alice,30,New York\r\n" +
            "Bob,25,London\r\n" +
            "Charlie,35,Sydney\r\n",
            Encoding.UTF8);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>, City: <<[row.City]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Enable reflection optimization for large data sets.
        ReportingEngine.UseReflectionOptimization = true;

        // Load CSV data source with headers.
        CsvDataSource csvData = new CsvDataSource(csvPath, new CsvDataLoadOptions { HasHeaders = true });

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "Data");

        // Save the generated report.
        doc.Save("output.docx");
    }
}
