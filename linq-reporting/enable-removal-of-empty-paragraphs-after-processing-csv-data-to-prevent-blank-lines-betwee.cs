using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // CsvDataSource, CsvDataLoadOptions, ReportingEngine, ReportBuildOptions are here.

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV handling.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare working folder.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // 1. Create sample CSV data.
            string csvPath = Path.Combine(workDir, "people.csv");
            File.WriteAllText(csvPath,
                "Name,Age,City\r\n" +
                "John Doe,30,New York\r\n" +
                "Jane Smith,,London\r\n" +          // Age empty.
                ",25,Paris\r\n" +                    // Name empty.
                "Bob Brown,45,\r\n");                // City empty.

            // 2. Build a template document with LINQ Reporting tags.
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin foreach over the CSV rows (exposed as a collection named "persons").
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("City: <<[person.City]>>");
            // End of foreach.
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // 3. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 4. Configure CSV data source.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
            {
                Delimiter = ',',
                QuoteChar = '"',
                CommentChar = '#',
                HasHeaders = true
            };
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // 5. Build the report with the option to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(reportDoc, csvDataSource, "persons");

            // 6. Save the final document.
            string outputPath = Path.Combine(workDir, "Report.docx");
            reportDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
