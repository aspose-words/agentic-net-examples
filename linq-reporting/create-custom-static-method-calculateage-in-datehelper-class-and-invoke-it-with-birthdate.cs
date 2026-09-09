using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Helper class containing the custom static method.
    public static class DateHelper
    {
        // Calculates age based on the provided birth date.
        public static int CalculateAge(DateTime birthDate)
        {
            var today = DateTime.Today;
            int age = today.Year - birthDate.Year;
            if (birthDate > today.AddYears(-age)) age--;
            return age;
        }
    }

    // Simple data model used as the root object for the report.
    public class Person
    {
        // Sample birth date property.
        public DateTime BirthDate { get; set; } = DateTime.MinValue;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that calls the static method.
            // The tag uses the root object's BirthDate property.
            builder.Writeln("Age: <<[DateHelper.CalculateAge(BirthDate)]>>");

            // Save the template locally.
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var person = new Person
            {
                // Example birth date.
                BirthDate = new DateTime(1990, 5, 15)
            };

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Register the helper type so its static members can be used in tags.
            engine.KnownTypes.Add(typeof(DateHelper));

            // Build the report using the root object name "person".
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "output.docx";
            doc.Save(outputPath);
        }
    }
}
