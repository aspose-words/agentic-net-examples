using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a simple data source – a list of objects.
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30, City = "New York" },
            new Person { Name = "Bob", Age = 25, City = "London" },
            new Person { Name = "Charlie", Age = 35, City = "Paris" }
        };

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a fixed number of columns (3 in this example).
        Table table = builder.StartTable();

        // ---- Header row ----
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.InsertCell();
        builder.Write("City");
        builder.EndRow();

        // ---- Data rows ----
        foreach (var person in people)
        {
            builder.InsertCell();
            builder.Write(person.Name);
            builder.InsertCell();
            builder.Write(person.Age.ToString());
            builder.InsertCell();
            builder.Write(person.City);
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "PeopleTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
    }

    // Simple POCO representing a row of data.
    private class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }
}
