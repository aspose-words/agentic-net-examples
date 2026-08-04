using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        string xmlPath = "persons.xml";
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // 1. Create a simple XML data source file.
        File.WriteAllText(xmlPath,
@"<persons>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </person>
</persons>");

        // 2. Build the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in persons]>>");

        // Start a table that will be repeated for each person.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row – values are filled by the reporting engine.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Age]>>");
        builder.EndRow();

        // Finish the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template and the XML data source.
        Document doc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 4. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // 5. Save the generated report.
        doc.Save(outputPath);
    }
}
