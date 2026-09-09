using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingPdfA
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for XML encoding support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file names.
            const string templatePath = "Template.docx";
            const string xmlDataPath = "Data.xml";
            const string outputPdfPath = "Report.pdf";

            // -----------------------------------------------------------------
            // 1. Create a simple XML data source file.
            // -----------------------------------------------------------------
            const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
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
            File.WriteAllText(xmlDataPath, xmlContent, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("==============");
            builder.Writeln();
            // LINQ Reporting foreach tag.
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the XML data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

            ReportingEngine engine = new ReportingEngine();
            // Build the report; the data source name must match the tag reference ("persons").
            engine.BuildReport(reportDoc, xmlDataSource, "persons");

            // -----------------------------------------------------------------
            // 4. Save the generated report as PDF/A compliant document.
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // PDF/A-1b compliance.
                Compliance = PdfCompliance.PdfA1b
            };
            reportDoc.Save(outputPdfPath, pdfOptions);
        }
    }
}
