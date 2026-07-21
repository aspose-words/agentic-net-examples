using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    // Simple data entity representing a record for the mail merge.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public string Message   { get; set; }

        public Person(string firstName, string lastName, string message)
        {
            FirstName = firstName;
            LastName  = lastName;
            Message   = message;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a mail‑merge template document in memory.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert merge fields that correspond to the Person properties.
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // 2. Define a collection of records to merge.
            List<Person> people = new List<Person>
            {
                new Person("John",  "Doe",   "Hello! This is your first merged document."),
                new Person("Jane",  "Smith", "Welcome to the mail merge example."),
                new Person("Bob",   "Brown", "Your order has been shipped.")
            };

            // 3. For each record, clone the template, execute the mail merge, and save a separate file.
            foreach (Person person in people)
            {
                // Clone the template to keep the original unchanged.
                Document doc = (Document)template.Clone();

                // Execute mail merge for a single record using field names and corresponding values.
                string[] fieldNames = { "FirstName", "LastName", "Message" };
                object[] fieldValues = { person.FirstName, person.LastName, person.Message };
                doc.MailMerge.Execute(fieldNames, fieldValues);

                // Build a unique filename for the merged document.
                string fileName = $"{person.FirstName}_{person.LastName}.docx";
                string filePath = Path.Combine(outputDir, fileName);

                // Save the merged document.
                doc.Save(filePath);
            }

            // The program finishes without waiting for user input.
        }
    }
}
