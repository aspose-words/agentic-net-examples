using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDemo
{
    // Simple data class that will be used as a LINQ data source for the ReportingEngine.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load a DOCX template that contains MailMerge fields and
            //    ReportingEngine tags.
            // -----------------------------------------------------------------
            // The template file must exist in the same folder as the executable
            // or provide a full path.
            string templatePath = "Template.docx";
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 2. Perform a simple MailMerge operation.
            // -----------------------------------------------------------------
            // Define the field names that exist in the template and the values
            // that will replace them.
            string[] mailMergeFieldNames = { "FullName", "Company", "Address", "City" };
            object[] mailMergeFieldValues = { "James Bond", "MI5 Headquarters", "Milbank", "London" };

            // Execute the mail merge. This fills the MERGEFIELD tags in the document.
            doc.MailMerge.Execute(mailMergeFieldNames, mailMergeFieldValues);

            // -----------------------------------------------------------------
            // 3. Populate the same document using the ReportingEngine with a LINQ
            //    data source (a List<Person> in this example).
            // -----------------------------------------------------------------
            // Create a collection of Person objects that will be referenced from
            // the template using the name "people".
            List<Person> people = new List<Person>
            {
                new Person { Name = "John Doe", Age = 28 },
                new Person { Name = "Jane Smith", Age = 34 },
                new Person { Name = "Bob Johnson", Age = 45 }
            };

            // Initialise the ReportingEngine.
            ReportingEngine reportingEngine = new ReportingEngine();

            // Build the report. The third parameter ("people") is the name that
            // will be used inside the template to reference the data source,
            // e.g. <<foreach [people]>><<[Name]>> <<[/foreach]>>.
            reportingEngine.BuildReport(doc, people, "people");

            // -----------------------------------------------------------------
            // 4. Save the resulting document to a new DOCX file.
            // -----------------------------------------------------------------
            string outputPath = "Result.docx";
            doc.Save(outputPath);

            // -----------------------------------------------------------------
            // 5. (Optional) Print the document to the default printer.
            // -----------------------------------------------------------------
            // Uncomment the following line if you want to send the document
            // directly to a printer.
            // doc.Print();

            Console.WriteLine($"Document generated and saved to '{outputPath}'.");
        }
    }
}
