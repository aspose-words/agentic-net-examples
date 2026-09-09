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

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath,
            "Name,Age,City" + Environment.NewLine +
            "Alice,30,New York" + Environment.NewLine +
            "Bob,25,London" + Environment.NewLine +
            "Charlie,35,Sydney");

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        // Begin foreach loop over CSV rows (named \"data\").
        builder.Writeln("<<foreach [record in data]>>");
        // Output each record's fields.
        builder.Writeln("Name: <<[record.Name]>>, Age: <<[record.Age]>>, City: <<[record.City]>>");
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure CSV load options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Create CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "data");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
