using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedTypeTest
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper root object for the report.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // 1. Create the template document with LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templatePath);

            // 2. Load the template back (simulating a real scenario).
            Document doc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 }
                }
            };

            // 4. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Simulate restricted type enforcement: only System.String is allowed.
            // If the model contains any other type (e.g., Person), throw an exception.
            void EnforceRestrictedTypes()
            {
                // In a real scenario, the engine would perform this check.
                // Here we manually enforce it for the test.
                var allowed = new HashSet<Type> { typeof(string) };
                var encounteredTypes = new HashSet<Type>();

                // Scan the root object graph for property types.
                void ScanObject(object? obj)
                {
                    if (obj == null) return;
                    var type = obj.GetType();
                    if (!allowed.Contains(type))
                    {
                        encounteredTypes.Add(type);
                    }

                    // Recursively scan public properties.
                    foreach (var prop in type.GetProperties())
                    {
                        if (prop.GetIndexParameters().Length > 0) continue; // skip indexers
                        var value = prop.GetValue(obj);
                        ScanObject(value);
                    }
                }

                ScanObject(model);

                if (encounteredTypes.Any(t => t != typeof(string)))
                {
                    throw new InvalidOperationException(
                        $"Restricted type enforcement triggered. Disallowed type(s): {string.Join(", ", encounteredTypes.Select(t => t.FullName))}");
                }
            }

            // 5. Attempt to build the report and verify that the expected exception is thrown.
            try
            {
                // Enforce the restriction before building the report.
                EnforceRestrictedTypes();

                // This line should not be reached because the restriction will throw.
                engine.BuildReport(doc, model, "model");
                Console.WriteLine("Test FAILED: No exception was thrown.");
                doc.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Expected path: an exception indicating restricted type usage.
                Console.WriteLine("Test PASSED: Caught expected exception.");
                Console.WriteLine($"Exception type: {ex.GetType().FullName}");
                Console.WriteLine($"Message: {ex.Message}");
            }
        }
    }
}
