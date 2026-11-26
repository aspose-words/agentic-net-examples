using System;
using System.IO;
using System.Reflection;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample class defined in this assembly.
    public class PersonInfo
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public PersonInfo Person { get; set; } = new();
        // Initialise with a simple XML element to avoid using a parameter‑less constructor.
        public XElement XmlData { get; set; } = new XElement("Root", "Default XML Value");
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "Template.docx");
            string reportPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Person Name: <<[model.Person.Name]>>");
            builder.Writeln("Person Age: <<[model.Person.Age]>>");
            builder.Writeln("XML Value: <<[model.XmlData.Value]>>");

            // Save the template.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. (Optional) Register external types.
            //    In recent versions of Aspose.Words this method is not required,
            //    and the API may not expose it. The engine can resolve types
            //    from the current assembly and System.Xml.Linq automatically.
            // -----------------------------------------------------------------
            // ReportingEngine.RegisterExternalTypes(new Assembly[]
            // {
            //     typeof(PersonInfo).Assembly,
            //     typeof(XElement).Assembly
            // });

            // -----------------------------------------------------------------
            // 4. Prepare the data model.
            // -----------------------------------------------------------------
            ReportModel model = new()
            {
                XmlData = new XElement("Root", "Sample XML Value")
            };

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
