using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the source PDF document.
        Document pdfDoc = new Document("Input.pdf");

        // Convert the Sections collection to a canonical array.
        // SectionCollection provides the ToArray() method.
        Section[] sectionsArray = pdfDoc.Sections.ToArray();

        // Load a template document that will be used for the report.
        Document templateDoc = new Document("Template.docx");

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the sections array as the data source.
        // The array can be passed directly as the data source object.
        engine.BuildReport(templateDoc, sectionsArray, "Sections");

        // Save the populated report to a DOCX file.
        templateDoc.Save("ReportFromPdf.docx");
    }
}
