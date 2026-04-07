using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a simple XML data source.
        string xmlPath = Path.Combine(workDir, "ReportData.xml");
        File.WriteAllText(xmlPath,
@"<Report>
    <Title>Monthly Sales Summary</Title>
    <Items>
        <Item>
            <Name>Product A</Name>
            <Quantity>120</Quantity>
        </Item>
        <Item>
            <Name>Product B</Name>
            <Quantity>85</Quantity>
        </Item>
        <Item>
            <Name>Product C</Name>
            <Quantity>47</Quantity>
        </Item>
    </Items>
</Report>");

        // 2. Build a Word template programmatically and insert LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report Title: <<[report.Title]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in report.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        string templatePath = Path.Combine(workDir, "Template.docx");
        template.Save(templatePath);

        // 3. Load the template and the XML data source.
        Document doc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are needed for this simple scenario.
        engine.BuildReport(doc, dataSource, "report");

        // 5. Save the result as a PDF/A‑2u compliant document.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u   // PDF/A‑2u compliance.
        };

        string outputPdf = Path.Combine(workDir, "Report.pdf");
        doc.Save(outputPdf, pdfOptions);

        // The example finishes here; the generated PDF/A file is located in the Work folder.
    }
}
