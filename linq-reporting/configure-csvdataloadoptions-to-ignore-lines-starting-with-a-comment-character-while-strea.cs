using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsCsvReporting
{
    // Simple data model representing a person.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for legacy encodings (required by Aspose.Words on .NET 5+).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file names in the current working directory.
            string csvPath = Path.Combine(Environment.CurrentDirectory, "people.csv");
            string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create a sample CSV file with comment lines.
            // -----------------------------------------------------------------
            // The comment character is '#'. Lines starting with this character will be ignored.
            string[] csvLines =
            {
                "# This is a comment line and should be skipped",
                "Name,Age",
                "Alice,30",
                "# Another comment that must be ignored",
                "Bob,45",
                "Charlie,25"
            };
            File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Create a Word template containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title.
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin a foreach loop over the CSV data source named "persons".
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template (simulating a separate load step).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Configure CSV loading options to ignore comment lines.
            // -----------------------------------------------------------------
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
            {
                Delimiter = ',',          // default delimiter, set explicitly for clarity
                CommentChar = '#',        // lines starting with '#' are ignored
                QuoteChar = '"'           // default quote character
            };

            // -----------------------------------------------------------------
            // 5. Create a CsvDataSource from the CSV file stream using the options.
            // -----------------------------------------------------------------
            using (FileStream csvStream = File.OpenRead(csvPath))
            {
                CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

                // -----------------------------------------------------------------
                // 6. Build the report using the ReportingEngine.
                // -----------------------------------------------------------------
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(loadedTemplate, dataSource, "persons");

                // -----------------------------------------------------------------
                // 7. Save the generated report.
                // -----------------------------------------------------------------
                loadedTemplate.Save(reportPath);
            }

            // The example finishes here. The generated report.docx contains the data
            // from the CSV file, with comment lines correctly ignored.
        }
    }
}
