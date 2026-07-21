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
            // Register code page provider for legacy encodings (required by Aspose.Words on .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths in the current working directory.
            string workDir = Directory.GetCurrentDirectory();
            string csvPath = Path.Combine(workDir, "people.csv");
            string templatePath = Path.Combine(workDir, "template.docx");
            string resultPath = Path.Combine(workDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create a sample CSV file with custom separator ';' and comment '#'.
            // -----------------------------------------------------------------
            string[] csvLines =
            {
                "# This is a comment line and will be ignored by the parser",
                "Name;Age;Comment",
                "Alice;30;First entry",
                "Bob;25;Second entry",
                "Charlie;35;Third entry"
            };
            File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Build a simple Word template containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Loop over the CSV rows (exposed as 'persons' data source).
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template, configure CSV load options, and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Configure CSV parsing: headers present, ';' as delimiter, '#' as comment character.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                Delimiter = ';',
                CommentChar = '#',
                HasHeaders = true
            };

            // Create the CSV data source using the configured options.
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // Build the report using the data source named "persons".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, csvDataSource, "persons");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(resultPath);

            // Optional: indicate completion (no interactive input).
            Console.WriteLine("Report generated successfully at: " + resultPath);
        }
    }
}
