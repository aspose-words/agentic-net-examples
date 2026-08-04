using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with comment lines
        string csvPath = "sample.csv";
        string[] csvLines =
        {
            "# This line is a comment and will be ignored",
            "Name,Age",
            "John,30",
            "# Another comment line",
            "Jane,25"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Create a simple template document containing LINQ Reporting tags
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template back (required by the workflow)
        Document loadedTemplate = new Document(templatePath);

        // Configure CSV loading options to ignore comment lines (starting with '#')
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
        {
            Delimiter = ',',
            CommentChar = '#',
            HasHeaders = true
        };

        // Create a CSV data source from the file stream using the configured options
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

            // Build the report using the data source; expose it to the template as "persons"
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, dataSource, "persons");
        }

        // Save the generated report
        string reportPath = "report.docx";
        loadedTemplate.Save(reportPath);
    }
}
