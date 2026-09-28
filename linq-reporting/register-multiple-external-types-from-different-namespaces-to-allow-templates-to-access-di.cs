using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // External namespace A
    namespace ExternalModels
    {
        public class Person
        {
            public string FirstName { get; set; } = "";
            public string LastName { get; set; } = "";

            // Computed full name used directly in the template.
            public string FullName => $"{FirstName} {LastName}";
        }
    }

    // External namespace B
    namespace ExternalData
    {
        public class Company
        {
            public string Name { get; set; } = "";

            // Upper‑cased name used directly in the template.
            public string UpperName => Name.ToUpperInvariant();
        }
    }

    // Root model used for the report
    public class ReportModel
    {
        public List<ExternalModels.Person> Persons { get; set; } = new();
        public ExternalData.Company Company { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a tag that uses the computed UpperName property.
            builder.Writeln("Company: <<[Company.UpperName]>>");

            // Begin a foreach loop over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("- <<[p.FullName]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a separate load step).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new()
            {
                Company = new ExternalData.Company { Name = "Acme Corp" },
                Persons = new()
                {
                    new ExternalModels.Person { FirstName = "John", LastName = "Doe" },
                    new ExternalModels.Person { FirstName = "Jane", LastName = "Smith" }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // No external type registration is required because the template
            // accesses instance properties (FullName, UpperName) directly.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "output.docx";
            doc.Save(outputPath);
        }
    }
}
