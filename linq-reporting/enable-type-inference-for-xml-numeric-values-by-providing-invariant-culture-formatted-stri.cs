using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare XML data with numeric values formatted using invariant culture.
        // Invariant culture ensures that numbers are represented with '.' as decimal separator,
        // which allows Aspose.Words LINQ Reporting to infer the correct numeric type.
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
    <person>
        <Name>John</Name>
        <Age>" + 30.ToString(CultureInfo.InvariantCulture) + @"</Age>
        <Salary>" + (12345.67m).ToString(CultureInfo.InvariantCulture) + @"</Salary>
    </person>
    <person>
        <Name>Jane</Name>
        <Age>" + 25.ToString(CultureInfo.InvariantCulture) + @"</Age>
        <Salary>" + (9876.54m).ToString(CultureInfo.InvariantCulture) + @"</Salary>
    </person>
</persons>";

        // Write the XML to a temporary file.
        string xmlPath = Path.Combine(Environment.CurrentDirectory, "persons.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Salary: <<[person.Salary]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report. The root object name must match the top‑level XML element ("persons").
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
