using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file names.
        const string xmlFile = "people.xml";
        const string templateFile = "template.docx";
        const string outputFile = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source file.
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<persons>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </person>
    <person>
        <Name>Bob Johnson</Name>
        <Age>40</Age>
    </person>
</persons>";
        File.WriteAllText(xmlFile, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk as required by the workflow.
        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 3. Load the template back and bind the XML data source.
        // -----------------------------------------------------------------
        var doc = new Document(templateFile);
        var xmlDataSource = new XmlDataSource(xmlFile);

        // Create the reporting engine. Do NOT enable AllowMissingMembers.
        var engine = new ReportingEngine();

        // Build the report using the data source name "persons".
        engine.BuildReport(doc, xmlDataSource, "persons");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputFile);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputFile)}");
    }
}
