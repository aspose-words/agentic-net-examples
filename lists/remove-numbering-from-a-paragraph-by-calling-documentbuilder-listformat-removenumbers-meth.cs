using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply default numbering to start a list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First item in the list");
        builder.Writeln("Second item in the list");

        // Remove numbering to stop list formatting for subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();
        builder.Writeln("This paragraph is not part of the list.");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
