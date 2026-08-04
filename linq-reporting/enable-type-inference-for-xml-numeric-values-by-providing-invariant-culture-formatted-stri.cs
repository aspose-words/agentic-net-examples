using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code page provider for XML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data with numeric values formatted using invariant culture.
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<People>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
        <Salary>1234.56</Salary>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>27</Age>
        <Salary>9876.54</Salary>
    </Person>
</People>";

        string xmlPath = "people.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Create a LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin the foreach block before the table.
        builder.Writeln("<<foreach [person in persons]>>");

        // Insert table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Salary");
        builder.EndRow();

        // Insert data row placeholders.
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Salary]>>");
        builder.EndRow();

        // Finish the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        var doc = new Document(templatePath);

        // 4. Load the XML data source using a stream.
        using (FileStream xmlStream = File.OpenRead(xmlPath))
        {
            var xmlDataSource = new XmlDataSource(xmlStream);
            var engine = new ReportingEngine();

            // Build the report. The root object name must match the tag reference ("persons").
            engine.BuildReport(doc, xmlDataSource, "persons");
        }

        // 5. Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
