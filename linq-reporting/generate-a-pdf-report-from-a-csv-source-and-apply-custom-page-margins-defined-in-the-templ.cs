using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample CSV data file.
        // -----------------------------------------------------------------
        const string csvFileName = "data.csv";
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,45",
            "Charlie,28"
        };
        File.WriteAllLines(csvFileName, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create a Word template with LINQ Reporting tags and custom margins.
        // -----------------------------------------------------------------
        const string templateFileName = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Apply custom page margins (1 inch = 72 points on each side).
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;

        // Add a title.
        builder.Writeln("Persons Report");
        builder.Writeln();

        // LINQ Reporting tags: iterate over the CSV rows (exposed as 'persons').
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age:  <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // -----------------------------------------------------------------
        // 3. Load the template (required by the rule set) and bind the CSV source.
        // -----------------------------------------------------------------
        Document doc = new Document(templateFileName);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Optional: specify delimiter, quote char, etc., if needed.
        // loadOptions.Delimiter = ',';
        // loadOptions.QuoteChar = '"';
        // loadOptions.CommentChar = '#';

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvFileName, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template tags is "persons".
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 4. Save the generated report as PDF.
        // -----------------------------------------------------------------
        const string outputPdf = "Report.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);
    }
}
