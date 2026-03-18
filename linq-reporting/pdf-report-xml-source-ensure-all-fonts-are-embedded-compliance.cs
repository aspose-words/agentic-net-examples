using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace PdfAReportGenerator
{
    public class ReportGenerator
    {
        /// <summary>
        /// Generates a PDF/A compliant report from an XML data source.
        /// </summary>
        /// <param name="templatePath">Path to the Word template document that contains the report tags.</param>
        /// <param name="xmlDataPath">Path to the XML file that provides the data for the report.</param>
        /// <param name="outputPdfPath">Path where the resulting PDF/A file will be saved.</param>
        public void GeneratePdfAReport(string templatePath, string xmlDataPath, string outputPdfPath)
        {
            // Load the Word template document.
            Document doc = new Document(templatePath);

            // Create an XML data source for the ReportingEngine.
            // XmlDataSource is part of Aspose.Words.Reporting and can be constructed from a file path.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

            // Populate the template with data from the XML source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, xmlDataSource);

            // Configure PDF save options for PDF/A compliance and full font embedding.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the compliance level to PDF/A-1b (visual appearance preservation).
            // Other levels such as PdfA2u or PdfA4 can be used as needed.
            saveOptions.Compliance = PdfCompliance.PdfA1b;

            // Ensure that all fonts are fully embedded (no subsetting) for compliance.
            saveOptions.EmbedFullFonts = true;                     // Embed the complete font files.
            saveOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll; // Embed every font used in the document.

            // Save the populated document as a PDF/A file.
            doc.Save(outputPdfPath, saveOptions);
        }
    }

    // Example usage:
    class Program
    {
        static void Main()
        {
            string templatePath = @"C:\Templates\ReportTemplate.docx";
            string xmlDataPath = @"C:\Data\ReportData.xml";
            string outputPdfPath = @"C:\Output\ReportPdfA.pdf";

            ReportGenerator generator = new ReportGenerator();
            generator.GeneratePdfAReport(templatePath, xmlDataPath, outputPdfPath);

            Console.WriteLine("PDF/A report generated successfully.");
        }
    }
}
