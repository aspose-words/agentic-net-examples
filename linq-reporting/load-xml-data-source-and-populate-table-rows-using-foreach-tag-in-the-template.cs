using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // File paths.
        const string templatePath = "template.docx";
        const string xmlPath = "data.xml";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source file.
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
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
        File.WriteAllText(xmlPath, xmlContent);

        // -----------------------------------------------------------------
        // 2. Build a Word template programmatically and embed LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Persons Report");
        builder.Writeln();

        // Begin the foreach block.
        builder.Writeln("<<foreach [person in persons]>>");

        // Create the table that will be repeated for each person.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row – will be repeated.
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and the XML data source, then build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        ReportingEngine engine = new ReportingEngine();
        // The data source name ("persons") must match the name used in the foreach tag.
        engine.BuildReport(reportDoc, dataSource, "persons");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
