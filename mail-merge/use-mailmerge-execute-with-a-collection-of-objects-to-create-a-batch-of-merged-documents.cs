using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear <<Name>>,");
        builder.Writeln("Your appointment is scheduled on <<Date>>.");
        builder.Writeln("Thank you.");

        // Prepare a collection of data objects.
        var data = new List<Person>
        {
            new Person { Name = "Alice", Date = DateTime.Today.AddDays(1).ToString("MMMM dd, yyyy") },
            new Person { Name = "Bob", Date = DateTime.Today.AddDays(2).ToString("MMMM dd, yyyy") },
            new Person { Name = "Charlie", Date = DateTime.Today.AddDays(3).ToString("MMMM dd, yyyy") }
        };

        // Perform mail merge for each object and save the merged document.
        int index = 1;
        foreach (var person in data)
        {
            // Clone the template for each merge operation.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge with the current object's data.
            mergedDoc.MailMerge.Execute(
                new[] { "Name", "Date" },
                new object[] { person.Name, person.Date });

            // Save the merged document to a file.
            string fileName = $"MergedDocument_{index}.docx";
            mergedDoc.Save(fileName);
            index++;
        }
    }

    // Simple class representing the data source for mail merge.
    public class Person
    {
        public string Name { get; set; }
        public string Date { get; set; }
    }
}
