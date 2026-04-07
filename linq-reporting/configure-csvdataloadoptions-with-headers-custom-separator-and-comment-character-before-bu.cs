using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsCsvReport
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV encoding support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths in the current working directory.
            string csvPath = Path.Combine(Environment.CurrentDirectory, "People.csv");
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a sample CSV file with headers, custom separator ';' and comment character '#'.
            // -----------------------------------------------------------------
            string[] csvLines =
            {
                "# This is a comment line and will be ignored",
                "Name;Age;Country",
                "Alice;30;USA",
                "Bob;25;Canada",
                "Charlie;35;UK"
            };
            File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Create a Word template containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("People Report");
            builder.Writeln();

            // Loop through each record in the CSV data source.
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("Country: <<[person.Country]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template document.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Configure CsvDataLoadOptions: headers present, ';' delimiter, '#' comment character.
            // -----------------------------------------------------------------
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                Delimiter = ';',
                CommentChar = '#'
            };

            // -----------------------------------------------------------------
            // 5. Create the CSV data source using the configured options.
            // -----------------------------------------------------------------
            CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

            // -----------------------------------------------------------------
            // 6. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The data source name "persons" matches the tag used in the template.
            engine.BuildReport(doc, dataSource, "persons");

            // -----------------------------------------------------------------
            // 7. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
