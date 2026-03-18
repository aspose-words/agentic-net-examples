using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that introduces the list.
        builder.Writeln("Key advantages of Aspose.Words:");

        // Start a default bulleted list. This creates a new list using the built‑in bullet template
        // and applies it to the current paragraph (the one just written above).
        builder.ListFormat.ApplyBulletDefault();

        // Add items that will be formatted with the default bullet.
        builder.Writeln("High performance");
        builder.Writeln("Rich and intuitive API");
        builder.Writeln("Cross‑platform support");
        builder.Writeln("Comprehensive documentation");

        // End the list. This removes the list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("BulletedList.docx");
    }
}
