using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the temporary files.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // 1. Create sample CSV data.
        File.WriteAllText(csvPath, "Name,Age\r\nAlice,30\r\nBob,25\r\nCharlie,35", Encoding.UTF8);

        // 2. Build the template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<foreach [in csvData]>>");
        builder.Writeln("<<[Name]>> - <<[Age]>>");
        builder.Writeln("<</foreach>>");
        template.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Prepare CSV data source with header support.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the data source.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, csvDataSource, "csvData");

        // 6. Save the generated report.
        reportDoc.Save(reportPath);
    }
}
