using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CustomTagDelimitersExample
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportData
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for older encodings (required by Aspose.Words on .NET 5+).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // 1. Create a template document that uses the default LINQ Reporting delimiters "<<", ">>".
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Correct foreach tag with proper delimiters.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>  Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a temporary file.
            string templatePath = Path.Combine(Path.GetTempPath(), "CustomTemplate.docx");
            template.Save(templatePath);

            // 2. Load the template back (simulating a real-world scenario where the template is read from disk).
            Document doc = new Document(templatePath);

            // 3. Prepare sample data.
            ReportData data = new ReportData
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob",   Age = 45 },
                    new Person { Name = "Carol", Age = 27 }
                }
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data");

            // 5. Save the generated report.
            string outputPath = Path.Combine(Path.GetTempPath(), "ReportResult.docx");
            doc.Save(outputPath);

            // Inform the user where the files are located.
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report generated at: {outputPath}");
        }
    }
}
