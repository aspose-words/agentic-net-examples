using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data source: a list of objects.
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a fixed number of columns (2 columns: Name and Age).
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Add a row for each item in the data source.
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

        // Save the document to a file.
        string outputPath = "TableOutput.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }

    // Simple data class representing a row in the table.
    private class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
