using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesSizeComparison
{
    // Simple data model used by the LINQ Reporting template.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper object that will be passed to the ReportingEngine.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 45 },
                    new Person { Name = "Charlie", Age = 28 }
                }
            };

            // Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a simple foreach tag that iterates over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a temporary file (required for the second run).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 1) Build report without InlineErrorMessages (default options).
            // -----------------------------------------------------------------
            var docWithoutInline = new Document(templatePath);
            var engineWithoutInline = new ReportingEngine(); // Options default to None.
            engineWithoutInline.BuildReport(docWithoutInline, model, "model");
            const string outputWithout = "Report_NoInlineError.docx";
            docWithoutInline.Save(outputWithout);

            // -----------------------------------------------------------------
            // 2) Build report with InlineErrorMessages enabled.
            // -----------------------------------------------------------------
            var docWithInline = new Document(templatePath);
            var engineWithInline = new ReportingEngine();
            engineWithInline.Options = ReportBuildOptions.InlineErrorMessages;
            bool success = engineWithInline.BuildReport(docWithInline, model, "model");
            // The success flag is meaningful only when InlineErrorMessages is set.
            const string outputWith = "Report_InlineError.docx";
            docWithInline.Save(outputWith);

            // -----------------------------------------------------------------
            // Compare file sizes.
            // -----------------------------------------------------------------
            long sizeWithout = new FileInfo(outputWithout).Length;
            long sizeWith = new FileInfo(outputWith).Length;

            Console.WriteLine($"Size without InlineErrorMessages: {sizeWithout} bytes");
            Console.WriteLine($"Size with InlineErrorMessages:    {sizeWith} bytes");
            Console.WriteLine($"Overhead introduced by InlineErrorMessages: {sizeWith - sizeWithout} bytes");
            Console.WriteLine($"BuildReport succeeded with InlineErrorMessages: {success}");
        }
    }
}
