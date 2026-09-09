using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingCaching
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[person.Name]>>!"); // LINQ Reporting tag.

            // Save the template to disk (required before building a report).
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Create the ReportingEngine instance.
            //    The engine automatically caches compiled templates internally,
            //    so no explicit UseCache property is required.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // -----------------------------------------------------------------
            // 4. Build the report the first time.
            // -----------------------------------------------------------------
            Person firstModel = new Person { Name = "Alice" };
            engine.BuildReport(loadedTemplate, firstModel, "person");
            const string firstReportPath = "Report1.docx";
            loadedTemplate.Save(firstReportPath);

            // -----------------------------------------------------------------
            // 5. Build the report a second time with different data.
            //    The engine reuses the compiled template from its internal cache.
            // -----------------------------------------------------------------
            Document secondRunTemplate = new Document(templatePath);
            Person secondModel = new Person { Name = "Bob" };
            engine.BuildReport(secondRunTemplate, secondModel, "person");
            const string secondReportPath = "Report2.docx";
            secondRunTemplate.Save(secondReportPath);
        }
    }
}
