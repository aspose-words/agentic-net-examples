using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingEnginePdfExample
{
    // Simple data class used as a data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int Age          { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template document that contains LINQ Reporting Engine tags.
            // Example tag in the template: <<[person.FirstName]>> <<[person.LastName]>> (Age: <<[person.Age]>>)
            string templatePath = @"C:\Templates\PersonReportTemplate.docx";

            // Path where the generated PDF will be saved.
            string outputPdfPath = @"C:\Output\PersonReport.pdf";

            // Load the template document from file.
            Document doc = new Document(templatePath);

            // Create a data source object. The name "person" will be used in the template.
            Person person = new Person
            {
                FirstName = "John",
                LastName  = "Doe",
                Age       = 42
            };

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by populating the template with the data source.
            // The third parameter is the name used to reference the data source inside the template.
            engine.BuildReport(doc, person, "person");

            // Save the populated document as PDF.
            // Using the overload that specifies the format ensures the correct file type.
            doc.Save(outputPdfPath, SaveFormat.Pdf);
        }
    }
}
