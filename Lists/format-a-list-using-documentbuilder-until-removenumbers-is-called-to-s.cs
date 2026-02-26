using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Path where the resulting DOCX will be saved.
        string outputPath = @"C:\Temp\BulletedList.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add list items – each Writeln call creates a new paragraph that becomes a list item.
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");
        builder.Writeln("Fourth bullet item");

        // End the bulleted list – removes bullets and resets the list level.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the specified location.
        doc.Save(outputPath);
    }
}
