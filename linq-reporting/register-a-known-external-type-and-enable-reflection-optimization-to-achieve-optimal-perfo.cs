using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Helper class whose static members can be used inside the template.
    public static class MyHelper
    {
        // Returns the input string in upper‑case.
        public static string Upper(string value) => value?.ToUpperInvariant() ?? string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Ensure the working directory exists.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(workDir);

            // 1. Create a sample XML file representing a large data set.
            string xmlPath = Path.Combine(workDir, "Persons.xml");
            CreateSampleXml(xmlPath, 5000); // 5 000 records for demonstration.

            // 2. Build a template document programmatically.
            string templatePath = Path.Combine(workDir, "Template.docx");
            CreateTemplateDocument(templatePath);

            // 3. Load the template (required before building the report).
            Document template = new Document(templatePath);

            // 4. Enable reflection optimization (static property).
            ReportingEngine.UseReflectionOptimization = true;

            // 5. Create the reporting engine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(MyHelper));

            // 6. Create an XML data source.
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // 7. Build the report. The data source name must match the tag used in the template.
            engine.BuildReport(template, dataSource, "persons");

            // 8. Save the generated report.
            string outputPath = Path.Combine(workDir, "ReportOutput.docx");
            template.Save(outputPath);
        }

        // Generates a simple XML file with the specified number of Person elements.
        private static void CreateSampleXml(string filePath, int count)
        {
            using StreamWriter writer = new StreamWriter(filePath);
            writer.WriteLine("<Persons>");
            for (int i = 1; i <= count; i++)
            {
                writer.WriteLine($"  <Person>");
                writer.WriteLine($"    <Name>Person {i}</Name>");
                writer.WriteLine($"    <Age>{20 + (i % 30)}</Age>");
                writer.WriteLine($"  </Person>");
            }
            writer.WriteLine("</Persons>");
        }

        // Creates a Word template containing LINQ Reporting tags.
        private static void CreateTemplateDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header.
            builder.Writeln("Persons Report");
            builder.Writeln("----------------");

            // Use a foreach loop over the XML data source named 'persons'.
            builder.Writeln("<<foreach [p in persons]>>");
            // Demonstrate usage of a registered external type.
            builder.Writeln("Name: <<[MyHelper.Upper(p.Name)]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
