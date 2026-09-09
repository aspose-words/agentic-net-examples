using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading before the list.
        builder.Writeln("Default numbered list:");

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the file system.
        doc.Save("NumberedList.docx");
    }
}
