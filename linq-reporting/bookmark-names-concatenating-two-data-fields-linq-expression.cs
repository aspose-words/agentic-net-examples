using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace AsposeWordsBookmarkDemo
{
    // Simple data entity with two fields that will be concatenated.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }

        public Person(string firstName, string lastName)
        {
            FirstName = firstName;
            LastName  = lastName;
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a sample data source.
            List<Person> people = new List<Person>
            {
                new Person("John", "Doe"),
                new Person("Jane", "Smith"),
                new Person("Bob",  "Johnson")
            };

            // Use LINQ to create a collection of concatenated bookmark names.
            // The format chosen here is "FirstName_LastName".
            var bookmarkNames = people
                .Select(p => $"{p.FirstName}_{p.LastName}")
                .ToList();

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph for each bookmark name and create the bookmark.
            foreach (string bmName in bookmarkNames)
            {
                // Start the bookmark with the generated name.
                builder.StartBookmark(bmName);
                // Add some placeholder text inside the bookmark.
                builder.Writeln($"Content for bookmark '{bmName}'.");
                // End the bookmark.
                builder.EndBookmark(bmName);
                // Add an empty line between entries for readability.
                builder.Writeln();
            }

            // Save the document to disk.
            doc.Save("BookmarksFromLinq.docx");
        }
    }
}
