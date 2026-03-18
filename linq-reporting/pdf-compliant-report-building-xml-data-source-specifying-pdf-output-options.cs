using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class PdfAReportGenerator
{
    static void Main()
    {
        // Define file paths (modify these paths to match your environment)
        string templatePath = @"C:\Docs\ReportTemplate.docx";
        string xmlDataPath = @"C:\Docs\ReportData.xml";
        string outputPath = @"C:\Docs\Report.pdf";

        // Load the Word template that contains reporting tags
        Document doc = new Document(templatePath);

        // Create an XML data source from the specified file
        XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

        // Populate the template with data using the ReportingEngine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "data");

        // Set up PDF save options to produce a PDF/A‑2u compliant document
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u   // PDF/A‑2u (ISO 19005‑2) compliance
        };

        // Save the filled report as a PDF/A document
        doc.Save(outputPath, saveOptions);
    }
}
