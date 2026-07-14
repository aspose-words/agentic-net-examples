using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare sample CSV data.
            const string csvFileName = "people.csv";
            File.WriteAllLines(csvFileName, new[]
            {
                "Name,Age,City",          // Header row
                "John Doe,30,New York",   // Normal row
                ",,,",                    // Empty row – will produce empty paragraph
                "Jane Smith,25,London"    // Another normal row
            });

            // Create a template document with LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the CSV rows (exposed as "persons").
            builder.Writeln("<<foreach [person in persons]>>");
            // Write each field on its own paragraph.
            builder.Writeln("<<[person.Name]>> - <<[person.Age]>> - <<[person.City]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Load the CSV data source with headers.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            CsvDataSource dataSource = new CsvDataSource(csvFileName, loadOptions);

            // Configure the reporting engine to remove empty paragraphs after processing.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the template and CSV data source.
            engine.BuildReport(doc, dataSource, "persons");

            // Save the generated report.
            const string outputFileName = "Report.docx";
            doc.Save(outputFileName);

            // Optional: indicate completion (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputFileName)}");
        }
    }
}
