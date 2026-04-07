using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple XML file with elements that have attributes.
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "items.xml");
        string xmlContent =
            @"<items>" +
            @"  <item Name=""Apple"" Color=""Red"" />" +
            @"  <item Name=""Banana"" Color=""Yellow"" />" +
            @"  <item Name=""Grape"" Color=""Purple"" />" +
            @"</items>";
        File.WriteAllText(xmlPath, xmlContent);

        // Create a Word template programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // LINQ Reporting foreach tag iterates over the XML elements.
        builder.Writeln("<<foreach [item in items]>>");
        // Concatenate attribute values using separate expression tags.
        builder.Writeln("<<[item.Name]>>-<<[item.Color]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report. The third argument is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine { Options = ReportBuildOptions.None };
        engine.BuildReport(doc, dataSource, "items");

        // Save the generated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);

        // Optional: display the path of the generated file.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
