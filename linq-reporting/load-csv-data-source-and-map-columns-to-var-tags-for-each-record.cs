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
        string csvPath = "people.csv";
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25\nCharlie,35");

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Configure CSV loading options (headers are present).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';
        loadOptions.HasHeaders = true;

        // Create CSV data source.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save("report.docx");
    }
}
