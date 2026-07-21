using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    // Simple data model for the table rows.
    private class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }

        public Person(string name, int age, string city)
        {
            Name = name;
            Age = age;
            City = city;
        }
    }

    public static void Main()
    {
        // Prepare a sample data source.
        List<Person> people = new List<Person>
        {
            new Person("Alice", 30, "New York"),
            new Person("Bob", 25, "London"),
            new Person("Charlie", 35, "Paris")
        };

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a fixed number of columns (3 in this case).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.InsertCell();
        builder.Write("City");
        builder.EndRow();

        // Data rows – added dynamically from the collection.
        foreach (Person p in people)
        {
            builder.InsertCell();
            builder.Write(p.Name);
            builder.InsertCell();
            builder.Write(p.Age.ToString());
            builder.InsertCell();
            builder.Write(p.City);
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicTable.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        // Optionally, inform that the process completed (no interactive prompts required).
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
