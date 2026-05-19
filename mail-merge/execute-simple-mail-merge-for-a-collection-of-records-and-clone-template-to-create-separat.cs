using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

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
            // Directory where the merged documents will be saved.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocs");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document that contains the MERGEFIELD tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // -----------------------------------------------------------------
            // 2. Prepare a collection of records to merge.
            // -----------------------------------------------------------------
            List<Person> people = new List<Person>
            {
                new Person("John",  "Doe",   "Hello! This is your first message."),
                new Person("Jane",  "Smith", "Welcome to the mail merge demo."),
                new Person("Bob",   "Brown", "Your order has been shipped.")
            };

            // -----------------------------------------------------------------
            // 3. For each record, clone the template, execute the mail merge,
            //    and save the result as a separate document.
            // -----------------------------------------------------------------
            int index = 1;
            foreach (Person person in people)
            {
                // Clone the template document. The Clone method returns an object,
                // so we cast it back to Document.
                Document mergedDoc = (Document)template.Clone(true);

                // Perform a simple mail merge for a single record.
                mergedDoc.MailMerge.Execute(
                    new[] { "FirstName", "LastName", "Message" },
                    new object[] { person.FirstName, person.LastName, person.Message });

                // Save the merged document.
                string fileName = Path.Combine(outputDir, $"Merged_{index}.docx");
                mergedDoc.Save(fileName);
                index++;
            }

            // The program finishes without waiting for user input.
        }
    }
}
