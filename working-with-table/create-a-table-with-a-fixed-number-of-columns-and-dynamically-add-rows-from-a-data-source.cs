using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Sample data source.
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // Create a new document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Start a table with a fixed number of columns (2).
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Dynamically add rows from the data source.
        foreach (var person in people)
        {
            builder.InsertCell();
            builder.Writeln(person.Name);
            builder.InsertCell();
            builder.Writeln(person.Age.ToString());
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "OutputTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Indicate success.
        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }

    // Simple data class representing a row.
    private class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
