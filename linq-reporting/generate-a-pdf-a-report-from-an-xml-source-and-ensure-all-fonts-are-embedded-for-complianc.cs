using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        File.WriteAllText(xmlPath,
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
    <Title>Sample PDF/A Report</Title>
    <Content>This is a sample report generated from XML data. All fonts will be embedded to meet PDF/A compliance.</Content>
</Report>");

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Writeln("<<[Report.Title]>>");
        builder.Writeln();
        builder.Writeln("<<[Report.Content]>>");
        templateDoc.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "Report");

        // Save as PDF/A with all fonts embedded.
        var pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            EmbedFullFonts = true,
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };
        string pdfPath = Path.Combine(outputDir, "report.pdf");
        doc.Save(pdfPath, pdfOptions);
    }
}
