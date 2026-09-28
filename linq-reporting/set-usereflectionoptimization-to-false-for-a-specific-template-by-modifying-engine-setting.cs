using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Output.docx";

            // -----------------------------------------------------------------
            // Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // LINQ Reporting tag.
            builder.Writeln("Hello <<[person.Name]>>!");
            templateDoc.Save(templatePath);

            // Load the template back from disk (required by the workflow).
            Document doc = new Document(templatePath);

            // Sample data to bind to the template.
            Person person = new Person { Name = "Aspose.Words" };

            // -----------------------------------------------------------------
            // Build the report with reflection optimization disabled.
            // -----------------------------------------------------------------
            // The static property controls whether the engine uses reflection
            // optimization. It must be set before creating the engine instance.
            ReportingEngine.UseReflectionOptimization = false;

            ReportingEngine engine = new ReportingEngine();
            // BuildReport returns void; no need for explicit disposal.
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
