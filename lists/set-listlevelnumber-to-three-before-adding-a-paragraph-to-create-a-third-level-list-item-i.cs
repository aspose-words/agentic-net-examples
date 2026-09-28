using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list to the document (using a numbered list template).
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Set the list level number to 3 to create a third‑level list item.
        builder.ListFormat.ListLevelNumber = 3;

        // Insert a paragraph as a third‑level list item.
        builder.Writeln("Third level list item");

        // Optional: clear list formatting after the item.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
