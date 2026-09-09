using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Enable code page provider for XML encoding support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a sample XML data file.
            const string xmlFileName = "People.xml";
            File.WriteAllText(xmlFileName,
@"<persons>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </person>
</persons>", Encoding.UTF8);

            // Build the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Heading.
            builder.Writeln("People Report");
            builder.Writeln();

            // LINQ Reporting tags.
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Optional: save the template for inspection.
            const string templateFileName = "ReportTemplate.docx";
            template.Save(templateFileName);

            // Load the XML data source.
            XmlDataSource dataSource = new XmlDataSource(xmlFileName);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "persons");

            // Save the generated report.
            const string outputFileName = "PeopleReport.docx";
            template.Save(outputFileName);

            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputFileName)}");
        }
    }
}
