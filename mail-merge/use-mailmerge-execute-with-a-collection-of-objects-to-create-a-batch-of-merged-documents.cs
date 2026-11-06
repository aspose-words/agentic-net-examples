using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");
        builder.Writeln("Thank you for your interest.");

        // Define a simple data source: an array of Person objects.
        Person[] people = new Person[]
        {
            new Person { FirstName = "John", LastName = "Doe" },
            new Person { FirstName = "Jane", LastName = "Smith" },
            new Person { FirstName = "Bob", LastName = "Johnson" }
        };

        // Field names must match the merge fields in the template.
        string[] fieldNames = new string[] { "FirstName", "LastName" };

        // Prepare output folder.
        string outputFolder = "MergedDocs";
        Directory.CreateDirectory(outputFolder);

        int docIndex = 1;
        foreach (Person person in people)
        {
            // Clone the template for each record.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge with the current object's values.
            object[] fieldValues = new object[] { person.FirstName, person.LastName };
            mergedDoc.MailMerge.Execute(fieldNames, fieldValues);

            // Validate that the merge replaced the fields.
            string mergedText = mergedDoc.GetText();
            if (string.IsNullOrWhiteSpace(mergedText) ||
                !mergedText.Contains(person.FirstName) ||
                !mergedText.Contains(person.LastName))
            {
                throw new InvalidOperationException($"Mail merge failed for record {docIndex}.");
            }

            // Save the merged document.
            string outputPath = Path.Combine(outputFolder, $"MergedDocument_{docIndex}.docx");
            mergedDoc.Save(outputPath);

            // Verify that the file was created and is not empty.
            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            {
                throw new InvalidOperationException($"Failed to create output file for record {docIndex}.");
            }

            docIndex++;
        }
    }

    // Simple POCO representing a data record.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
    }
}
