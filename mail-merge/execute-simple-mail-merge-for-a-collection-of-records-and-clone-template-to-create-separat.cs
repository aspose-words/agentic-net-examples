using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a template document that contains merge fields.
        Document template = CreateTemplate();

        // Sample data records.
        List<Person> people = new List<Person>
        {
            new Person("John", "Doe", "Hello John! This is a merged message."),
            new Person("Jane", "Smith", "Welcome Jane! Your document is ready."),
            new Person("Bob", "Johnson", "Good day Bob! Here is your content.")
        };

        // Names of the merge fields used in the template.
        string[] fieldNames = { "FirstName", "LastName", "Message" };

        // Iterate over each record, clone the template, perform mail merge, and save.
        int index = 1;
        foreach (Person person in people)
        {
            // Clone the template document.
            Document doc = (Document)template.Clone();

            // Values for the current record.
            object[] fieldValues = { person.FirstName, person.LastName, person.Message };

            // Execute mail merge for a single record.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document.
            string filePath = Path.Combine(outputDir, $"MergedDocument_{index}.docx");
            doc.Save(filePath);
            index++;
        }
    }

    // Creates a simple template document with MERGEFIELDs.
    private static Document CreateTemplate()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        return doc;
    }

    // Simple data class representing a record for mail merge.
    private class Person
    {
        public string FirstName { get; }
        public string LastName { get; }
        public string Message { get; }

        public Person(string firstName, string lastName, string message)
        {
            FirstName = firstName;
            LastName = lastName;
            Message = message;
        }
    }
}
