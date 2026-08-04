using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample XML data source.
        string xmlPath = Path.Combine(outputDir, "persons.xml");
        string xmlContent =
@"<Persons>
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
</Persons>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        // Insert the current date/time directly; no need for a reporting tag.
        builder.Writeln($"Generated on: {DateTime.Now}");
        builder.Writeln(); // empty line

        // Begin foreach loop over the collection named "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Load XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // 6. Save the result as PDF/A‑1b.
        string pdfPath = Path.Combine(outputDir, "PersonsReport.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };
        doc.Save(pdfPath, pdfOptions);

        Console.WriteLine($"Report generated successfully: {pdfPath}");
    }
}
