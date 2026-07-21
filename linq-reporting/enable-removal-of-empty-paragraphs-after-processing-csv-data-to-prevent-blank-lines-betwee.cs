using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingCsv
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV parsing (required for non‑UTF8 encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample CSV data.
            string csvPath = "people.csv";
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Age,City",
                "John Doe,30,New York",
                "Jane Smith,25,London",
                ",,",
                "Bob Johnson,40,Paris"
            });

            // Create a template document with LINQ Reporting tags.
            string templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            Document doc = new Document(templatePath);

            // Configure CSV data source options.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
            {
                Delimiter = ',',
                CommentChar = '#',
                QuoteChar = '"'
            };
            CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

            // Build the report with the option to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(doc, dataSource, "persons");

            // Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        // Creates a simple Word template containing a foreach loop over CSV rows.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header.
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin foreach loop over the CSV rows (named "persons").
            builder.Writeln("<<foreach [person in persons]>>");

            // Paragraphs that will be populated with data.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("City: <<[person.City]>>");

            // An extra empty paragraph that may become empty if all fields are empty.
            builder.Writeln();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
