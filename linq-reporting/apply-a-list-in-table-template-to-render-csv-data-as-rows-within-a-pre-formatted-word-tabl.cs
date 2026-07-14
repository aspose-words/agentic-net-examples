using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing on .NET Core.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        string csvPath = "people.csv";
        string[] csvLines =
        {
            "Name,Age,City",
            "Alice,30,London",
            "Bob,25,Paris",
            "Charlie,35,New York"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document with a pre‑formatted table and LINQ
        //    Reporting tags that will iterate over the CSV rows.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Title (optional)
        builder.Writeln("People Report");
        builder.Writeln();

        // Start the foreach block – iterate over the data source named "data".
        builder.Writeln("<<foreach [person in data]>>");

        // Create a table with a header row.
        var table = builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Data row – each cell contains a tag that outputs a column value.
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.City]>>");
        builder.EndRow();

        // End of the table and foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save("Template.docx");

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using the CSV data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document("Template.docx");

        // Configure CSV loading options – the file has a header row.
        var loadOptions = new CsvDataLoadOptions(true)
        {
            HasHeaders = true,
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };

        // Create the CSV data source.
        var csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the final report.
        reportDoc.Save("Report.docx");
    }
}
