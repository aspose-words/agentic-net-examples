using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class PdfAReportGenerator
{
    static void Main()
    {
        // Create a simple Word document with a reporting tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<? for-each $data.item ?>");
        builder.Writeln("Name: <? $item.Name ?>");
        builder.Writeln("Age: <? $item.Age ?>");
        builder.Writeln("<? end for-each ?>");

        // Sample XML data as a string.
        string xml = @"
<data>
    <item>
        <Name>John Doe</Name>
        <Age>30</Age>
    </item>
    <item>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </item>
</data>";

        // Load the XML data from a memory stream.
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        XmlDataSource dataSource = new XmlDataSource(xmlStream);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "data");

        // Set PDF/A‑2u compliance options.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the resulting PDF in the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.pdf");
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
