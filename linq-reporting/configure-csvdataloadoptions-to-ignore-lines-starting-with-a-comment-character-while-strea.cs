using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        string templatePath = "Template.docx";
        string csvPath = "Data.csv";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a LINQ Reporting template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple foreach loop that will iterate over the CSV rows.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a sample CSV file with comment lines.
        // -----------------------------------------------------------------
        // The CSV has a header row, a comment line (starting with '#'), and two data rows.
        string[] csvLines =
        {
            "Name,Age",
            "# This line is a comment and should be ignored",
            "Alice,30",
            "Bob,25"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Configure CsvDataLoadOptions to ignore comment lines.
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
        {
            Delimiter = ',',      // default delimiter, set explicitly for clarity
            CommentChar = '#',    // lines starting with '#' will be ignored
            QuoteChar = '"'       // default quote character
        };

        // -----------------------------------------------------------------
        // 4. Load the CSV data as a stream and create a CsvDataSource.
        // -----------------------------------------------------------------
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

            // Load the previously saved template document.
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 5. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "persons");

            // Save the generated report.
            doc.Save(reportPath);
        }

        // The example finishes without waiting for user input.
    }
}
