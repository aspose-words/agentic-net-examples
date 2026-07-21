using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class PdfAReportExample
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample XML data.
        string xmlPath = Path.Combine(workDir, "people.xml");
        File.WriteAllText(xmlPath,
@"<root>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
    <Person>
        <Name>Bob Johnson</Name>
        <Age>40</Age>
    </Person>
</root>");

        // 2. Build a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Persons Report");
        builder.Writeln();

        // Since the XML root element directly contains a list of Person nodes,
        // the root object itself is the collection. Use it in the foreach tag.
        builder.Writeln("<<foreach [person in root]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // 3. Load the template for report generation.
        Document report = new Document(templatePath);

        // 4. Create an XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        // The root object name used in the template tags is "root".
        engine.BuildReport(report, dataSource, "root");

        // 6. Save the populated document as PDF/A‑1b.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };
        string outputPdf = Path.Combine(workDir, "PersonsReport.pdf");
        report.Save(outputPdf, pdfOptions);

        Console.WriteLine($"Report generated: {outputPdf}");
    }
}
