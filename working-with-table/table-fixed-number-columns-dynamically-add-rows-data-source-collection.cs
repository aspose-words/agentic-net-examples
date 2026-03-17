using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableDemo
{
    // Simple data class representing a row in the table.
    public class Person
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

    public static class TableBuilder
    {
        // Creates a Word document containing a table with a fixed number of columns (3)
        // and adds one row for each element in the supplied collection.
        public static Document CreateDocumentWithTable(IEnumerable<Person> data)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new table and add it to the document body.
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Ensure the table has at least one row and one cell so we can start adding content.
            table.EnsureMinimum();

            // ---------- Header row ----------
            Row headerRow = new Row(doc);
            table.AppendChild(headerRow);

            // Helper to create a header cell with given text.
            void AddHeaderCell(string text)
            {
                Cell cell = new Cell(doc);
                cell.AppendChild(new Paragraph(doc));
                cell.FirstParagraph.AppendChild(new Run(doc, text));
                // Make header text bold.
                cell.FirstParagraph.Runs[0].Font.Bold = true;
                headerRow.AppendChild(cell);
            }

            AddHeaderCell("Name");
            AddHeaderCell("Age");
            AddHeaderCell("City");

            // ---------- Data rows ----------
            foreach (var person in data)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                // Name cell
                Cell nameCell = new Cell(doc);
                nameCell.AppendChild(new Paragraph(doc));
                nameCell.FirstParagraph.AppendChild(new Run(doc, person.Name));
                row.AppendChild(nameCell);

                // Age cell
                Cell ageCell = new Cell(doc);
                ageCell.AppendChild(new Paragraph(doc));
                ageCell.FirstParagraph.AppendChild(new Run(doc, person.Age.ToString()));
                row.AppendChild(ageCell);

                // City cell
                Cell cityCell = new Cell(doc);
                cityCell.AppendChild(new Paragraph(doc));
                cityCell.FirstParagraph.AppendChild(new Run(doc, person.City));
                row.AppendChild(cityCell);
            }

            // Auto‑fit the table to its contents for a tidy layout.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            return doc;
        }

        // Example usage.
        public static void Main()
        {
            // Sample data source.
            List<Person> people = new List<Person>
            {
                new Person("Alice", 30, "New York"),
                new Person("Bob",   25, "London"),
                new Person("Carol", 28, "Paris")
            };

            // Build the document.
            Document doc = CreateDocumentWithTable(people);

            // Save the document to disk.
            doc.Save("TableWithDynamicRows.docx");
        }
    }
}
