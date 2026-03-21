using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class PdfReportGenerator
{
    static void Main()
    {
        // Paths to the template, CSV data source and the output PDF file.
        string templatePath = "Template.docx";
        string csvPath = "Data.csv";
        string outputPath = "Report.pdf";

        // Ensure a simple template exists.
        if (!File.Exists(templatePath))
        {
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Report");
            builder.Writeln("Name: <<data.Name>>");
            builder.Writeln("Age: <<data.Age>>");
            templateDoc.Save(templatePath);
        }

        // Ensure a simple CSV data source exists.
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath, "Name,Age\nJohn Doe,30");
        }

        // Load the template document.
        Document doc = new Document(templatePath);

        // Apply custom page margins to the first section.
        var pageSetup = doc.Sections[0].PageSetup;
        pageSetup.TopMargin = ConvertUtil.InchToPoint(0.5);      // 0.5 inch top margin
        pageSetup.BottomMargin = ConvertUtil.InchToPoint(0.5);   // 0.5 inch bottom margin
        pageSetup.LeftMargin = ConvertUtil.InchToPoint(0.75);    // 0.75 inch left margin
        pageSetup.RightMargin = ConvertUtil.InchToPoint(0.75);   // 0.75 inch right margin

        // Configure CSV loading options (first line contains column headers).
        var loadOptions = new CsvDataLoadOptions(true);

        // Create a CSV data source from the file.
        var csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report by merging the CSV data into the template.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "data");

        // Save the populated document as a PDF file.
        var pdfOptions = new PdfSaveOptions();
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF report generated: {Path.GetFullPath(outputPath)}");
    }
}
