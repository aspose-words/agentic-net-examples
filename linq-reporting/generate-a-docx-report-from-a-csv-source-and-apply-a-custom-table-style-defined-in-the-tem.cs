using System;
using System.IO;
using System.Text;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample CSV data file.
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        File.WriteAllText(csvPath,
            "Product,Quantity,Price\r\n" +
            "Apple,10,0.5\r\n" +
            "Banana,5,0.3\r\n" +
            "Carrot,12,0.2\r\n");

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Define a custom table style.
        TableStyle customStyle = (TableStyle)template.Styles.Add(StyleType.Table, "CustomTableStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 10;
        customStyle.Shading.BackgroundPatternColor = Color.LightGray;
        customStyle.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.Color = Color.DarkGray;
        customStyle.Borders.LineWidth = 0.5;

        // Begin the foreach block.
        builder.Writeln("<<foreach [row in data]>>");

        // Start the table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data row – each cell contains a LINQ Reporting tag.
        builder.InsertCell();
        builder.Writeln("<<[row.Product]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Price]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Apply the custom style to the table.
        table.StyleName = "CustomTableStyle";

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using CSV data source.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // Configure CSV load options – the file has a header row.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);

        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root object name used in the template tags is "data".
        engine.BuildReport(report, csvData, "data");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }
}
