using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV reading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Quantity,Price\n" +
            "1,Apple,10,0.5\n" +
            "2,Banana,5,0.3\n" +
            "3,Cherry,20,0.2");

        // Create a Word template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Set custom page margins (1 inch on each side).
        builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1);
        builder.PageSetup.RightMargin = ConvertUtil.InchToPoint(1);
        builder.PageSetup.TopMargin = ConvertUtil.InchToPoint(1);
        builder.PageSetup.BottomMargin = ConvertUtil.InchToPoint(1);

        // Add a title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Writeln("Product Report");
        builder.Font.Size = 12;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln();

        // Insert LINQ Reporting tags to iterate over CSV rows.
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Quantity: <<[row.Quantity]>>");
        builder.Writeln("Price: $<<[row.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a file (optional, for inspection).
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template back (demonstrates the required load step).
        Document doc = new Document(templatePath);

        // Configure CSV data source.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
            // Default separator is ',' and default encoding is UTF8, which match our CSV.
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "data");

        // Save the generated report as PDF.
        string outputPdf = "Report.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);
    }
}
