using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file names in the current working directory.
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "persons.xml");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportOutput.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source file.
        // -----------------------------------------------------------------
        string xmlContent =
@"<persons>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </person>
</persons>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection of persons.
        builder.Writeln("<<foreach [p in persons]>>");
        // Output each person's name and age.
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using the XML data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        ReportingEngine engine = new ReportingEngine();
        // The data source name ("persons") must match the name used in the template tags.
        engine.BuildReport(reportDoc, dataSource, "persons");

        // -----------------------------------------------------------------
        // 4. Save the generated report to a DOCX file.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
