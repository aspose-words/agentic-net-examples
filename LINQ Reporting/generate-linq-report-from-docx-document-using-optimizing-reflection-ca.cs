using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Sample data class used as a LINQ data source.
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
            // Path to the template DOCX file that contains LINQ Reporting Engine tags,
            // e.g. <<foreach [in ds]>><<[FirstName]>><<[LastName]>> (<[Age]>)<</foreach>>
            string templatePath = @"C:\Templates\PeopleReport.docx";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\PeopleReport_Output.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a collection of Person objects as the data source.
            List<Person> people = new List<Person>
            {
                new Person("John",  "Doe",      30),
                new Person("Jane",  "Smith",    25),
                new Person("Alice", "Johnson",  28)
            };

            // -----------------------------------------------------------------
            // Optimize reflection calls.
            // By default Aspose.Words uses dynamic class generation to speed up
            // reflection. For small collections the overhead of generating the
            // dynamic class may outweigh the benefit, so we can disable it.
            // -----------------------------------------------------------------
            // Enable optimization (default behavior).
            ReportingEngine.UseReflectionOptimization = true;

            // If you know the data set is small, you may disable it like this:
            // ReportingEngine.UseReflectionOptimization = false;

            // Create the reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument ("ds") is the name used in the
            // template to reference the data source object itself.
            engine.BuildReport(doc, people, "ds");

            // Save the populated document.
            doc.Save(outputPath);
        }
    }
}
