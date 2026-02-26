using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingToXps
{
    // Simple POCO class that will be used as the data source for the LINQ Reporting Engine.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int Age          { get; set; }

        // Example of an expression that can be evaluated in the template.
        public string FullName => $"{FirstName} {LastName}";
    }

    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains LINQ Reporting tags, e.g. <<[person.FullName]>>.
            const string templatePath = @"C:\Templates\PersonReport.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a data source instance.
            Person person = new Person
            {
                FirstName = "John",
                LastName  = "Doe",
                Age       = 42
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name used inside the template to reference the data source.
            engine.BuildReport(doc, person, "person");

            // Prepare XPS save options (default constructor is allowed by the rules).
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Save the populated document as XPS.
            const string outputPath = @"C:\Output\PersonReport.xps";
            doc.Save(outputPath, xpsOptions);
        }
    }
}
