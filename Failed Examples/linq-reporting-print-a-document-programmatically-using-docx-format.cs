// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsPrintExample
{
    // Simple data class used as a data source for the reporting engine.
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
            // Path to the DOCX template that contains reporting tags, e.g. <<[person.FirstName]>>.
            string templatePath = @"C:\Templates\PersonReport.docx";

            // Load the template document (lifecycle rule: load).
            Document doc = new Document(templatePath);

            // Prepare the data source.
            Person person = new Person
            {
                FirstName = "John",
                LastName  = "Doe",
                Age       = 42
            };

            // Build the report by merging the data source with the template (LINQ reporting).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // Optional: save the populated document to a new file (lifecycle rule: save).
            string outputPath = @"C:\Output\PersonReport_Filled.docx";
            doc.Save(outputPath); // SaveFormat is inferred from the .docx extension.

            // Print the document to the default printer (printing rule).
            doc.Print();

            // If you need to print to a specific printer, uncomment the following line and set the printer name.
            // doc.Print("Your Printer Name");
        }
    }
}
