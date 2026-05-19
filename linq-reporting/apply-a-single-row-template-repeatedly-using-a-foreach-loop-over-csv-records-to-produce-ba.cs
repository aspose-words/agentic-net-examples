using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBatch
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV parsing (required on .NET Core).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string workDir = Directory.GetCurrentDirectory();
            string dataDir = Path.Combine(workDir, "Data");
            string outputDir = Path.Combine(workDir, "Output");
            Directory.CreateDirectory(dataDir);
            Directory.CreateDirectory(outputDir);

            // 1. Create a sample CSV file.
            string csvPath = Path.Combine(dataDir, "people.csv");
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Age",
                "Alice,30",
                "Bob,45",
                "Charlie,28"
            });

            // 2. Create a single‑row template document programmatically.
            string templatePath = Path.Combine(dataDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // 3. Load the template.
            Document reportDoc = new Document(templatePath);

            // 4. Configure CSV data source options.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // first line has headers
            loadOptions.Delimiter = ','; // default delimiter
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // 5. Build the report using LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(reportDoc, csvDataSource, "persons");

            // 6. Save the generated batch document.
            string outputPath = Path.Combine(outputDir, "PeopleReport.docx");
            reportDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine("Report generated at: " + outputPath);
        }
    }
}
