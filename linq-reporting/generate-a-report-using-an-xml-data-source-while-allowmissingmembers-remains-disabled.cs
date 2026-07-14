using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Ensure the CodePages provider is registered (required for some XML encodings).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Paths for the files used in the example.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string xmlPath = Path.Combine(outputDir, "people.xml");
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");

            // 1. Create a simple XML data source file.
            // The XML contains a root element <persons> with multiple <person> entries.
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

            // 2. Build a Word template programmatically.
            // The template uses LINQ Reporting tags to iterate over the XML data.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the collection named "persons".
            builder.Writeln("<<foreach [person in persons]>>");
            // Output each person's name and age.
            builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 3. Load the template document for reporting.
            Document loadedTemplate = new Document(templatePath);

            // 4. Create an XmlDataSource from the XML file.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // 5. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Do NOT enable AllowMissingMembers; the default (None) keeps the option disabled.
            engine.Options = ReportBuildOptions.None;

            // 6. Build the report.
            // The root object name in the template is "persons", matching the XML root element.
            engine.BuildReport(loadedTemplate, xmlDataSource, "persons");

            // 7. Save the generated report.
            loadedTemplate.Save(reportPath);

            // Inform the user where the files are located.
            Console.WriteLine("Template, XML data source, and report have been created in:");
            Console.WriteLine(outputDir);
        }
    }
}
