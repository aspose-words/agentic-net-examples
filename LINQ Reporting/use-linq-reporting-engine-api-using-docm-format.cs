using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data source class used in the template.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains LINQ Reporting Engine tags.
            Document doc = new Document("Template.docm");

            // Prepare the data that will be merged into the template.
            var person = new Person
            {
                FirstName = "John",
                LastName  = "Doe",
                Age       = 30
            };

            // Create an instance of the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument is the name used to reference the data source in the template.
            engine.BuildReport(doc, person, "person");

            // Save the populated document as a macro‑enabled DOCM file.
            doc.Save("Report.docm");
        }
    }
}
