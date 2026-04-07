using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        const string xmlFile = "Data.xml";
        const string templateFile = "ReportTemplate.docx";
        const string outputFile = "ReportResult.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source file.
        // -----------------------------------------------------------------
        string xmlContent =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Persons>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
</Persons>";
        File.WriteAllText(xmlFile, xmlContent);

        // -----------------------------------------------------------------
        // 2. Build a Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        // Loop over the collection of Person elements.
        // The data source will be referenced by the name "persons".
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age : <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 3. Load the template, bind the XML data source and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templateFile);
        XmlDataSource dataSource = new XmlDataSource(xmlFile);

        ReportingEngine engine = new ReportingEngine();
        // Pass the data source name ("persons") so that the template can reference it.
        engine.BuildReport(reportDoc, dataSource, "persons");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputFile);
    }
}
