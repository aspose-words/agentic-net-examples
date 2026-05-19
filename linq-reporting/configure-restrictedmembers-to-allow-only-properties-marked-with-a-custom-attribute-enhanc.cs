using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace RestrictedMembersExample
{
    // Custom attribute to mark properties that are allowed to be accessed in the template.
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExposeAttribute : Attribute { }

    // Sample data model with some properties marked with the custom attribute.
    public class Person
    {
        [Expose] public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        [Expose] public string Email { get; set; } = string.Empty;
        public string SecretNote { get; set; } = string.Empty;
    }

    // Wrapper that only exposes properties marked with [Expose].
    public class PersonReport
    {
        public string Name { get; }
        public string Email { get; }

        public PersonReport(Person source)
        {
            Name = source.Name;
            Email = source.Email;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some data sources).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // 1. Create a simple Word template with LINQ Reporting tags.
            var templatePath = "template.docx";
            CreateTemplate(templatePath);

            // 2. Prepare sample data.
            var person = new Person
            {
                Name = "John Doe",
                Age = 42,
                Email = "john.doe@example.com",
                SecretNote = "This should not be visible in the report."
            };

            // 3. Restrict the original Person type so its members cannot be accessed directly.
            //    This forces the engine to use the wrapper that only contains exposed members.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // 4. Build the report using the wrapper as the root data source.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine
            {
                // Allow missing members so that attempts to access non‑exposed properties do not throw.
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = string.Empty
            };

            var wrapper = new PersonReport(person);
            engine.BuildReport(doc, wrapper, "model");

            // 5. Save the generated report.
            doc.Save("Report.docx");
        }

        // Helper method to create a Word document containing the required tags.
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Report");
            builder.Writeln("----------------");
            builder.Writeln("Name : <<[model.Name]>>");
            builder.Writeln("Email: <<[model.Email]>>");
            // This line attempts to access a non‑exposed property; it will be ignored.
            builder.Writeln("Age  : <<[model.Age]>>");

            doc.Save(filePath);
        }
    }
}
