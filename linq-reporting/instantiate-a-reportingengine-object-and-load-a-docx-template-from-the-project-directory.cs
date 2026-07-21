using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model used by the template.
    public class Person
    {
        public Person(string name)
        {
            Name = name;
        }

        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Define file names relative to the project directory.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Step 1: Create a DOCX template with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple paragraph containing a tag that references the data model.
            builder.Writeln("Hello, <<[person.Name]>>!");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template from the file system.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Prepare the data source.
            // -----------------------------------------------------------------
            Person person = new Person("Aspose");

            // -----------------------------------------------------------------
            // Step 4: Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "person".
            engine.BuildReport(loadedTemplate, person, "person");

            // -----------------------------------------------------------------
            // Step 5: Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }
}
