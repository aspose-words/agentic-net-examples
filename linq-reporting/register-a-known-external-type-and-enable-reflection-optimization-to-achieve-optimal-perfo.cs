using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Helper class with a static method that will be accessed from the template.
    public static class Utils
    {
        public static string Upper(string value) => value?.ToUpperInvariant() ?? string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for XML handling.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample XML data.
            const string xmlFileName = "people.xml";
            File.WriteAllText(xmlFileName,
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
                        <Age>45</Age>
                    </person>
                </persons>");

            // Create a template document with LINQ Reporting tags.
            const string templateFileName = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            // Use a static member from a known external type (Math.PI) to demonstrate registration.
            builder.Writeln("Pi: <<[Math.PI]>>");
            // Use a static method from our custom Utils class (dot syntax works for static calls).
            builder.Writeln("Upper Name: <<[Utils.Upper(p.Name)]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templateFileName);

            // Load the template for reporting.
            var doc = new Document(templateFileName);

            // Enable reflection optimization for large data sets.
            ReportingEngine.UseReflectionOptimization = true;

            // Create the reporting engine and register known external types.
            var engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Math));   // System.Math for static members.
            engine.KnownTypes.Add(typeof(Utils)); // Custom utility class.

            // Load XML data source.
            var xmlDataSource = new XmlDataSource(File.OpenRead(xmlFileName));

            // Build the report. The root object name is "persons" as used in the template.
            engine.BuildReport(doc, xmlDataSource, "persons");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
