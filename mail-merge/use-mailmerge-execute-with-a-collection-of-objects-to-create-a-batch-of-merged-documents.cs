using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");
        builder.Writeln("This is a test mail merge document.");

        // Prepare a collection of data objects.
        List<Person> people = new List<Person>
        {
            new Person { FirstName = "John", LastName = "Doe" },
            new Person { FirstName = "Jane", LastName = "Smith" },
            new Person { FirstName = "Bob", LastName = "Johnson" }
        };

        // Define merge field names.
        string[] fieldNames = new string[] { "FirstName", "LastName" };

        // Ensure output directory exists.
        string outputDir = "MergedDocs";
        Directory.CreateDirectory(outputDir);

        int index = 1;
        foreach (Person person in people)
        {
            // Clone the template for each record.
            Document doc = (Document)template.Clone(true);

            // Prepare field values for the current record.
            object[] fieldValues = new object[] { person.FirstName, person.LastName };

            // Execute the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Validate that the merge replaced the fields.
            string mergedText = doc.GetText();
            if (!mergedText.Contains(person.FirstName) || !mergedText.Contains(person.LastName))
                throw new InvalidOperationException($"Mail merge failed for record {index}.");

            // Save the merged document.
            string filePath = Path.Combine(outputDir, $"Merged_{index}.docx");
            doc.Save(filePath);
            index++;
        }
    }

    // Simple data class representing a person.
    private class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
    }
}
