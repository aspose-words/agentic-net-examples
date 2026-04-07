using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingXmlExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the files used in the example.
            const string xmlFilePath = "people.xml";
            const string templatePath = "template.docx";
            const string outputPath = "report.docx";

            // 1. Create a simple XML data source file.
            // The root element is <persons> containing multiple <person> entries.
            string xmlContent =
                @"<?xml version=""1.0"" encoding=""utf-8""?>
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
            File.WriteAllText(xmlFilePath, xmlContent);

            // 2. Build a Word template programmatically and embed LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 3. Load the template back (simulating a real‑world scenario where the template is a file).
            Document loadedTemplate = new Document(templatePath);

            // 4. Create an XmlDataSource that reads the XML file created earlier.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlFilePath);

            // 5. Use ReportingEngine to populate the template.
            // Do NOT enable AllowMissingMembers – the default options are used.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the root name used in the template tags ("persons").
            engine.BuildReport(loadedTemplate, xmlDataSource, "persons");

            // 6. Save the generated report.
            loadedTemplate.Save(outputPath);

            // Inform the user that the process completed.
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
