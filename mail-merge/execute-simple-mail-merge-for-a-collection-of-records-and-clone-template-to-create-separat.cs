using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document that contains the merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare a collection of records to merge.
        var people = new List<Person>
        {
            new Person("John", "Doe", "Hello! This is the first message."),
            new Person("Jane", "Smith", "Greetings from the second record."),
            new Person("Bob", "Johnson", "Another message for you.")
        };

        // For each record, clone the template, execute the mail merge, and save a separate document.
        int docIndex = 1;
        foreach (var person in people)
        {
            // Clone the template to keep the original unchanged.
            Document mergedDoc = (Document)template.Clone();

            // Perform a simple mail merge for a single record.
            mergedDoc.MailMerge.Execute(
                new[] { "FirstName", "LastName", "Message" },
                new object[] { person.FirstName, person.LastName, person.Message });

            // Save the merged document with a unique filename.
            string outputFileName = $"MergedDocument_{docIndex}.docx";
            mergedDoc.Save(outputFileName);
            docIndex++;
        }
    }

    // Simple data class representing a record for the mail merge.
    private class Person
    {
        public Person(string firstName, string lastName, string message)
        {
            FirstName = firstName;
            LastName = lastName;
            Message = message;
        }

        public string FirstName { get; }
        public string LastName { get; }
        public string Message { get; }
    }
}
