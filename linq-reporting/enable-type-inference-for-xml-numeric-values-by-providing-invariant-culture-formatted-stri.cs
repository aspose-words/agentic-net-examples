using System;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Step 1: Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a foreach loop over the XML data source named "persons".
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Step 2: Create sample XML data with numeric values formatted using invariant culture.
        var xmlPath = "Data.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>27</Age>
    </person>
    <person>
        <Name>Bob Johnson</Name>
        <Age>45</Age>
    </person>
</persons>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Step 3: Load the template document.
        var doc = new Document(templatePath);

        // Step 4: Create an XmlDataSource from the XML file.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Step 5: Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // Ensure numeric values are parsed using invariant culture.
        doc.FieldOptions.UseInvariantCultureNumberFormat = true;
        engine.BuildReport(doc, xmlDataSource, "persons");

        // Step 6: Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
