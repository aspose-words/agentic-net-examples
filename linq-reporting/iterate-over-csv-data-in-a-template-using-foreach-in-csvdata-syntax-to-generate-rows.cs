using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string workDir = Directory.GetCurrentDirectory();
        string dataPath = Path.Combine(workDir, "data.csv");
        string templatePath = Path.Combine(workDir, "Template.docx");
        string reportPath = Path.Combine(workDir, "Report.docx");

        // Create sample CSV file.
        File.WriteAllText(dataPath,
            "Name,Age,Country\n" +
            "Alice,30,USA\n" +
            "Bob,25,Canada\n" +
            "Charlie,35,UK");

        // Build the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over CSV data.
        builder.Writeln("<<foreach [row in csvData]>>");

        // Start table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Country");
        builder.EndRow();

        // Data row – use LINQ Reporting tags.
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Country]>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
            // Default separator is ','; no need to set explicitly.
        };
        CsvDataSource csvData = new CsvDataSource(dataPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "csvData");

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}
