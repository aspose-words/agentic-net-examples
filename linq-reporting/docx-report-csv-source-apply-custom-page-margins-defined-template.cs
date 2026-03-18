using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportGenerator
{
    static void Main()
    {
        // Paths to the template, CSV source and the output document.
        string templatePath = @"C:\Docs\ReportTemplate.docx";
        string csvPath = @"C:\Docs\SourceData.csv";
        string outputPath = @"C:\Docs\GeneratedReport.docx";

        // Load the DOCX template (create → load).
        Document reportDoc = new Document(templatePath);

        // Create a CSV data source (load).
        CsvDataSource csvData = new CsvDataSource(csvPath);

        // Build the report by merging the CSV data into the template.
        // The third parameter is the name used to reference the data source inside the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "Data");

        // Apply custom page margins.
        // Set the preset to Custom and then define the individual margins in points (1 inch = 72 points).
        if (reportDoc.Sections.Count > 0)
        {
            var pageSetup = reportDoc.Sections[0].PageSetup;
            pageSetup.Margins = Margins.Custom;          // Use custom margins.
            pageSetup.TopMargin = 72;    // 1.0 inch
            pageSetup.BottomMargin = 72; // 1.0 inch
            pageSetup.LeftMargin = 108;  // 1.5 inches
            pageSetup.RightMargin = 108; // 1.5 inches
        }

        // Save the populated document (save).
        reportDoc.Save(outputPath);
    }
}
