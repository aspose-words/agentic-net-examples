using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for Table type

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        });

        // -----------------------------------------------------------------
        // 2. Build a template document that contains a foreach block.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Open foreach block – the data source will be referenced as 'persons'.
        builder.Writeln("<<foreach [person in persons]>>");

        // Create a two‑column table inside the foreach block.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.EndRow();
        builder.EndTable();

        // Close foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // CSV options: first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "persons");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save("Report.docx");
    }
}
