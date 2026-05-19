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

        // Paths for temporary files.
        string templatePath = "Template.docx";
        string csvPath = "Data.csv";
        string outputPath = "Report.docx";

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a heading.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin a foreach block over the CSV rows (named "persons").
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Active: <<[person.IsActive]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk, then reload it as required by the lifecycle rule.
        templateDoc.Save(templatePath);
        Document loadedTemplate = new Document(templatePath);

        // ---------- Create sample CSV data ----------
        // The CSV has a header row and a boolean column represented as true/false strings.
        string[] csvLines =
        {
            "Name,IsActive",
            "Alice,true",
            "Bob,false",
            "Charlie,true"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // ---------- Configure CSV load options ----------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line contains headers
        {
            Delimiter = ',',   // default delimiter
            QuoteChar = '"',   // default quote character
            CommentChar = '#', // any comment lines start with '#'
        };

        // ---------- Create CSV data source ----------
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // ---------- Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        // The data source name "persons" matches the tag used in the template.
        engine.BuildReport(loadedTemplate, csvDataSource, "persons");

        // Save the generated report.
        loadedTemplate.Save(outputPath);
    }
}
