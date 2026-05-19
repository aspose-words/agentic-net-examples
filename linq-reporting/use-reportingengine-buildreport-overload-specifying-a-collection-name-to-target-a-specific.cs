using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper object that contains the collection.
    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new Model();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(new Person { Name = "Bob", Age = 45 });
            model.Persons.Add(new Person { Name = "Charlie", Age = 28 });

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Add a heading.
            builder.Writeln("Person Report");
            builder.Writeln();

            // foreach tag iterating over the collection exposed by the wrapper object.
            builder.Writeln("<<foreach [p in model.Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report using the overload that
            //    specifies the data source name ("model").
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);
            var engine = new ReportingEngine();

            // The third argument ("model") allows the template to reference the wrapper object.
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
