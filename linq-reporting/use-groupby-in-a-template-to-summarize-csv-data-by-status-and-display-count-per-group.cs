using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current directory.
        string csvPath = Path.Combine(Environment.CurrentDirectory, "data.csv");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        // Columns: Id,Name,Status
        string[] csvLines =
        {
            "Id,Name,Status",
            "1,Alpha,Open",
            "2,Beta,Closed",
            "3,Gamma,Open",
            "4,Delta,InProgress",
            "5,Epsilon,Closed",
            "6,Zeta,Open"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Status Summary Report");
        builder.Writeln();

        // Loop over groups of persons by Status and display the count.
        builder.Writeln("<<foreach [g in persons.GroupBy(p => p.Status)]>>");
        builder.Writeln("Status: <<[g.Key]>>  -  Count: <<[g.Count()]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure CSV loading options (first row contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Create a CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "persons".
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
