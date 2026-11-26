using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    // Simple data model for rows
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    public static void Main()
    {
        // Sample data source
        List<Person> people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30, City = "New York" },
            new Person { Name = "Bob", Age = 25, City = "London" },
            new Person { Name = "Charlie", Age = 35, City = "Paris" }
        };

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional title
        builder.Writeln("Dynamic Table Example");

        // Start the table with a fixed number of columns (3)
        builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Add a row for each person in the collection
        foreach (var person in people)
        {
            builder.InsertCell();
            builder.Writeln(person.Name);
            builder.InsertCell();
            builder.Writeln(person.Age.ToString());
            builder.InsertCell();
            builder.Writeln(person.City);
            builder.EndRow();
        }

        // End the table
        builder.EndTable();

        // Save the document
        string outputPath = "TableOutput.docx";
        doc.Save(outputPath);

        // Verify that the file was created
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
