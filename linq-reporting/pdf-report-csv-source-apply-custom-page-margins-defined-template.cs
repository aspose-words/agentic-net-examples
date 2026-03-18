using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Reporting; // for CsvDataSource and CsvDataLoadOptions

class PdfReportGenerator
{
    static void Main()
    {
        // Paths to the template, CSV data source and the output PDF file.
        string templatePath = "Template.docx";
        string csvPath = "Data.csv";
        string outputPath = "Report.pdf";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Apply custom page margins to the first (and usually only) section.
        // Use the preset 'Custom' and then set the individual margins in points.
        // 1 inch = 72 points; ConvertUtil.InchToPoint can be used for conversion.
        doc.Sections[0].PageSetup.Margins = Margins.Custom;
        doc.Sections[0].PageSetup.TopMargin = ConvertUtil.InchToPoint(0.5);    // 0.5 inch top margin
        doc.Sections[0].PageSetup.BottomMargin = ConvertUtil.InchToPoint(0.5); // 0.5 inch bottom margin
        doc.Sections[0].PageSetup.LeftMargin = ConvertUtil.InchToPoint(0.75);  // 0.75 inch left margin
        doc.Sections[0].PageSetup.RightMargin = ConvertUtil.InchToPoint(0.75); // 0.75 inch right margin

        // Configure CSV loading options (assume the first line contains column headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Optional: customize delimiter, comment character, etc. if needed.
        // loadOptions.Delimiter = ',';
        // loadOptions.CommentChar = '#';

        // Create a CSV data source from the file.
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report by merging the CSV data into the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "data");

        // Prepare PDF save options (default options are sufficient for this scenario).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the populated document as a PDF file.
        doc.Save(outputPath, pdfOptions);
    }
}
