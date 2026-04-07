using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple XML data source file.
        const string xmlFileName = "people.xml";
        File.WriteAllText(xmlFileName,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
  <person>
    <Name>John Doe</Name>
    <Age>30</Age>
  </person>
  <person>
    <Name>Jane Smith</Name>
    <Age>25</Age>
  </person>
</persons>");

        // Build a Word template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Use a common font that will be embedded later.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;

        builder.Writeln("People Report");
        builder.Writeln();

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required by the lifecycle rule).
        const string templateFileName = "template.docx";
        template.Save(templateFileName);

        // Load the template back for report generation.
        Document report = new Document(templateFileName);

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlFileName);

        // Populate the template with data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, dataSource, "persons");

        // Configure PDF/A save options with full font embedding.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1b;
        pdfOptions.EmbedFullFonts = true;

        // Save the final report as PDF/A.
        report.Save("PeopleReport.pdf", pdfOptions);
    }
}
