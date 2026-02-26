using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class used in the template.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int Age          { get; set; }

        public Person(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName  = lastName;
            Age       = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains LINQ Reporting tags,
            // e.g. <<[person.FirstName]>> <<[person.Age]>>
            string templatePath = @"Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a data source instance.
            Person person = new Person("John", "Doe", 42);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("person") must match the name used in the template tags.
            engine.BuildReport(doc, person, "person");

            // Configure HTML Fixed save options.
            HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
            {
                SaveFormat = SaveFormat.HtmlFixed, // Ensure the format is HTML Fixed.
                ExportEmbeddedImages = true,       // Embed images directly into the HTML.
                ShowPageBorder = false,            // Optional: hide page borders in the output.
                PrettyFormat = true                // Optional: make the HTML more readable.
            };

            // Path for the resulting HTML Fixed file.
            string outputPath = @"Report.html";

            // Save the populated document as HTML Fixed.
            doc.Save(outputPath, htmlOptions);
        }
    }
}
