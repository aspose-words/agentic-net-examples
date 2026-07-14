using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class PdfAReportGenerator
{
    public static void Main()
    {
        // Register code page provider for XML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string xmlPath = Path.Combine(outputDir, "data.xml");
        string pdfPath = Path.Combine(outputDir, "report.pdf");

        // 1. Create sample XML data.
        string xmlContent =
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Report>
    <Title>Sample PDF/A Report</Title>
    <Date>2023-01-01</Date>
    <Items>
        <Item>
            <Name>Widget A</Name>
            <Quantity>10</Quantity>
        </Item>
        <Item>
            <Name>Widget B</Name>
            <Quantity>25</Quantity>
        </Item>
        <Item>
            <Name>Widget C</Name>
            <Quantity>7</Quantity>
        </Item>
    </Items>
</Report>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // 2. Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title and date.
        builder.Writeln("<<[report.Title]>>");
        builder.Writeln("Date: <<[report.Date]>>");
        builder.Writeln();

        // Items list using foreach tag.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in report.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 4. Load XML data source.
        using (FileStream xmlStream = File.OpenRead(xmlPath))
        {
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // 5. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(reportDoc, dataSource, "report");
        }

        // 6. Configure PDF/A save options with full font embedding.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b, // PDF/A-1b compliance.
            EmbedFullFonts = true,             // Embed every glyph of every font.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll // Ensure all fonts are embedded.
        };

        // 7. Save the final report as PDF/A.
        reportDoc.Save(pdfPath, pdfOptions);
    }
}
