using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create a sample CSV file.
        const string csvPath = "sample.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Product,Quantity,Price",
            "Apple,10,0.5",
            "Banana,5,0.3",
            "Carrot,7,0.2"
        });

        // 2. Build the template document programmatically.
        const string templatePath = "template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a title.
        builder.Writeln("Report generated from CSV data");
        builder.Writeln();

        // Define a custom table style named "MyTableStyle".
        // Cast the added style to TableStyle to access table‑specific formatting.
        TableStyle tableStyle = (TableStyle)template.Styles.Add(StyleType.Table, "MyTableStyle");
        tableStyle.Shading.BackgroundPatternColor = Color.LightGray;          // Light gray header background.
        tableStyle.Borders.Color = Color.DarkGray;                           // Dark gray border color.
        tableStyle.Borders.LineWidth = 1.0;                                  // Border thickness.

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [row in data]>>");

        // Start the table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row – values will be filled from the CSV source.
        builder.InsertCell();
        builder.Writeln("<<[row.Product]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Price]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Apply the custom style to the table.
        table.StyleName = "MyTableStyle";
        table.AutoFit(AutoFitBehavior.AutoFitToContents);
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the template.
        template.Save(templatePath);

        // 3. Load the template for reporting.
        Document report = new Document(templatePath);

        // 4. Prepare the CSV data source.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // first line contains headers
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, csvData, "data");

        // 6. Save the final report.
        const string outputPath = "ReportFromCsv.docx";
        report.Save(outputPath);

        // Inform that the process completed.
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
