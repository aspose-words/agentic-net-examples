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
            // Ensure the working directory is the executable's directory.
            string workDir = AppDomain.CurrentDomain.BaseDirectory;

            // 1. Create a sample XML data file.
            string xmlPath = Path.Combine(workDir, "people.xml");
            File.WriteAllText(xmlPath,
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
</persons>");

            // 2. Build a template document that contains LINQ Reporting tags.
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Create an XmlDataSource from the XML file.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // 5. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(doc, xmlDataSource, "persons");

            // 6. Save the generated report.
            string outputPath = Path.Combine(workDir, "output.docx");
            doc.Save(outputPath);

            // Inform that the process completed (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
