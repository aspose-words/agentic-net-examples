using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file names in the current working directory.
        string csvPath = Path.Combine(Environment.CurrentDirectory, "data.csv");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportOutput.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers and numeric values.
        // -----------------------------------------------------------------
        File.WriteAllText(csvPath,
            "Name,Value1,Value2\r\n" +
            "Item A,10,15\r\n" +
            "Item B,7,3\r\n" +
            "Item C,20,5\r\n",
            Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template containing LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== LINQ Reporting Example ===");
        builder.Writeln("<<foreach [row in persons]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Value1: <<[row.Value1]>>");
        builder.Writeln("Value2: <<[row.Value2]>>");
        // Calculated field: sum of Value1 and Value2.
        builder.Writeln("Sum: <<[row.Value1 + row.Value2]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Configure CSV data source options (headers present, comma delimiter).
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            CommentChar = '#', // No comments in our file, but set a default.
            QuoteChar = '"'
        };

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 5. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options required for this simple example.
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
