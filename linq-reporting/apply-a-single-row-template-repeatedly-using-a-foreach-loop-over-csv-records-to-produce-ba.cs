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
        string workDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workDir, "people.csv");
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // Create a simple CSV file with headers.
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        });

        // Build a template document containing a foreach tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Start the foreach block over the CSV rows (named "persons").
        builder.Writeln("<<foreach [person in persons]>>");
        // Output each person's data on a separate line.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back for reporting.
        Document doc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';
        loadOptions.HasHeaders = true;

        // Create a CSV data source from the file.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
