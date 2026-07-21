using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the working directory exists.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // Paths for the template, XML data, and output report.
            string templatePath = Path.Combine(workDir, "Template.docx");
            string xmlPath = Path.Combine(workDir, "Data.xml");
            string reportPath = Path.Combine(workDir, "Report.docx");

            // 1. Create a simple Word template with LINQ Reporting tags.
            // The template references a missing XML element (LastName) to demonstrate handling.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("First Name: <<[person.FirstName]>>");
            builder.Writeln("Last Name: <<[person.LastName]>>"); // This element is missing in the XML.
            templateDoc.Save(templatePath);

            // 2. Create an XML file where the <LastName> element is intentionally omitted.
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <person>
        <FirstName>John</FirstName>
        <!-- LastName element is missing -->
    </person>
</root>";
            File.WriteAllText(xmlPath, xmlContent);

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Load the XML data source.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // 5. Configure the ReportingEngine to treat missing members as empty strings.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Optional: customize the message shown for missing members (empty string suppresses output).
            engine.MissingMemberMessage = string.Empty;

            // 6. Build the report. The empty string for dataSourceName means members are accessed directly.
            engine.BuildReport(doc, xmlDataSource, "");

            // 7. Save the generated report.
            doc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine("Report generated successfully at:");
            Console.WriteLine(reportPath);
        }
    }
}
