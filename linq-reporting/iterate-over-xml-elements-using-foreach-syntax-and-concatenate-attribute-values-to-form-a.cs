using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model is not required because we use XmlDataSource directly.
    class Program
    {
        static void Main()
        {
            // Prepare sample XML data with attributes.
            const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<items>
    <item name=""Apple"" value=""10"" />
    <item name=""Banana"" value=""20"" />
    <item name=""Cherry"" value=""30"" />
</items>";
            const string xmlPath = "data.xml";
            File.WriteAllText(xmlPath, xmlContent);

            // Create a template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            // The foreach tag iterates over each <item> element.
            // Inside the loop we output the attribute values concatenated with a space.
            builder.Writeln("<<foreach [item in items]>>");
            builder.Writeln("<<[item.name]>> <<[item.value]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // Load XML data source from the file.
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "items", matching the top‑level XML element.
            engine.BuildReport(reportDoc, dataSource, "items");

            // Save the generated report.
            const string outputPath = "output.docx";
            reportDoc.Save(outputPath);

            Console.WriteLine("Report generated successfully: " + Path.GetFullPath(outputPath));
        }
    }
}
