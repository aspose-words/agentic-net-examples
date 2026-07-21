using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingExample
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required in some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output documents
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document with LINQ tags
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Persons Report");
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare the data source
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportModel model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(new Person { Name = "Bob", Age = 45 });
            model.Persons.Add(new Person { Name = "Charlie", Age = 28 });

            // -------------------------------------------------
            // 3. Build the report using ReportingEngine
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options required for this simple example
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated report
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
