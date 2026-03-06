// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a DOCX template that contains MERGEFIELDs and ReportingEngine tags.
        Document doc = new Document("Template.docx");

        // ---------- Mail Merge ----------
        // Simple mail merge using an array of field names and values.
        string[] fieldNames = { "Name", "Address" };
        object[] fieldValues = { "John Doe", "123 Main St." };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // ---------- Reporting Engine ----------
        // Create a LINQ data source (list of Person objects).
        var people = new List<Person>
        {
            new Person { Name = "Alice",   Age = 30 },
            new Person { Name = "Bob",     Age = 45 },
            new Person { Name = "Charlie", Age = 25 }
        };

        // Example LINQ query: select only adults (Age >= 30).
        var adults = people.Where(p => p.Age >= 30).ToList();

        // Populate the template tags (e.g., <<[ds.Name]>>) using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name "ds" is used inside the template to reference the collection.
        engine.BuildReport(doc, adults, "ds");

        // ---------- Print ----------
        // Send the document to the default printer.
        doc.Print();

        // ---------- Save ----------
        // Save the final document to a DOCX file.
        doc.Save("Result.docx");
    }

    // Simple POCO used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
