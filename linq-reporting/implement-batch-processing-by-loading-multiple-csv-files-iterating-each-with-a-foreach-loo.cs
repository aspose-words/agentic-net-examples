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

        // Create a folder for sample CSV files.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        // Generate two sample CSV files.
        CreateSampleCsv(Path.Combine(dataFolder, "people1.csv"));
        CreateSampleCsv(Path.Combine(dataFolder, "people2.csv"));

        // Create a template document that contains LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        CreateTemplate(templatePath);

        // Document that will hold the merged results from all CSV files.
        Document masterDocument = new Document();

        // Process each CSV file in the folder.
        foreach (string csvFile in Directory.GetFiles(dataFolder, "*.csv"))
        {
            // Load the template for the current CSV file.
            Document templateDocument = new Document(templatePath);

            // Configure CSV loading options (headers are present).
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                HasHeaders = true
            };

            // Create a CSV data source from the file.
            CsvDataSource csvData = new CsvDataSource(csvFile, loadOptions);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(templateDocument, csvData, "persons");

            // Append the generated report to the master document.
            masterDocument.AppendDocument(templateDocument, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedReport.docx");
        masterDocument.Save(outputPath);
    }

    // Generates a simple CSV file with Name and Age columns.
    private static void CreateSampleCsv(string filePath)
    {
        string[] lines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25"
        };
        File.WriteAllLines(filePath, lines);
    }

    // Creates a Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("People Report");
        builder.Writeln("----------------");

        // Begin foreach loop over the CSV rows (exposed as 'persons').
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>\tAge: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}
