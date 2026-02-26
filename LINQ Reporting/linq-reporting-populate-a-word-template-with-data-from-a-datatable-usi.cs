using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains MERGEFIELDs matching the column names.
        string templatePath = "Template.dotx";
        Document doc = new Document(templatePath);

        // Sample data source: a list of Person objects.
        List<Person> people = new List<Person>
        {
            new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
            new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
            new Person { FirstName = "Alice", LastName = "Brown", Age = 28 }
        };

        // Create a DataTable and define its schema.
        DataTable table = new DataTable("People");
        table.Columns.Add("FirstName", typeof(string));
        table.Columns.Add("LastName",  typeof(string));
        table.Columns.Add("Age",       typeof(int));

        // Populate the DataTable using LINQ projection.
        var query = people.Select(p => new { p.FirstName, p.LastName, p.Age });
        foreach (var row in query)
        {
            table.Rows.Add(row.FirstName, row.LastName, row.Age);
        }

        // Execute mail merge: fill the template with data from the DataTable.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Simple POCO representing a data entity.
    class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }
    }
}
