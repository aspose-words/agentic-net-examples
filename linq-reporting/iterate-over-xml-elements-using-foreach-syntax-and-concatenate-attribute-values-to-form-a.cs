using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // File names for the template, data source and the generated report.
        const string templatePath = "Template.docx";
        const string xmlDataPath = "Data.xml";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source with element values.
        // -----------------------------------------------------------------
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>"
          + "<persons>"
          + "  <person>"
          + "    <Name>John</Name>"
          + "    <Age>30</Age>"
          + "  </person>"
          + "  <person>"
          + "    <Name>Anna</Name>"
          + "    <Age>25</Age>"
          + "  </person>"
          + "  <person>"
          + "    <Name>Mike</Name>"
          + "    <Age>40</Age>"
          + "  </person>"
          + "</persons>";
        File.WriteAllText(xmlDataPath, xmlContent);

        // -----------------------------------------------------------------
        // 2. Build a Word template that uses LINQ Reporting foreach tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach over the collection 'persons'.
        builder.Writeln("<<foreach [p in persons]>>");
        // Concatenate the element values 'Name' and 'Age' with a hyphen.
        builder.Writeln("<<[p.Name]>>-<<[p.Age]>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and the XML data source.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The data source name must match the name used in the template tags.
        engine.BuildReport(loadedTemplate, xmlDataSource, "persons");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);

        // Inform the user where the report was saved.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
