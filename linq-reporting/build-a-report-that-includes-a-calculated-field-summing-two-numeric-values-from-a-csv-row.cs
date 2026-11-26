using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Item,Value1,Value2\n" +
            "Apple,10,5\n" +
            "Banana,7,3\n" +
            "Cherry,12,8");

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over CSV rows.
        builder.Writeln("<<foreach [row in rows]>>");

        // Start the table inside the loop.
        Table table = builder.StartTable();

        // Header row (only once, so we add a condition to render it on the first iteration).
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Sum");
        builder.EndRow();

        // Data row – first cell.
        builder.InsertCell();
        builder.Writeln("<<[row.Item]>>");

        // Data row – second cell (calculated sum).
        builder.InsertCell();
        builder.Writeln("<<[row.Value1 + row.Value2]>>");

        // End of the data row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Configure CSV data source (default separator is comma).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "rows");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
