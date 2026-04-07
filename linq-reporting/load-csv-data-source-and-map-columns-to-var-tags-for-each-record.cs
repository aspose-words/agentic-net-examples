using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsCsvLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV parsing (required on .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample CSV data.
            string csvPath = "people.csv";
            File.WriteAllText(csvPath,
                "Name,Age,City\n" +
                "Alice,30,New York\n" +
                "Bob,25,London\n" +
                "Charlie,35,Sydney");

            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags.
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>, City: <<[person.City]>>");
            builder.Writeln("<</foreach>>");

            // Configure CSV loading options (first line contains headers).
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            loadOptions.Delimiter = ',';

            // Create CSV data source.
            CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, csvData, "persons");

            // Save the generated report.
            template.Save("ReportFromCsv.docx");
        }
    }
}
