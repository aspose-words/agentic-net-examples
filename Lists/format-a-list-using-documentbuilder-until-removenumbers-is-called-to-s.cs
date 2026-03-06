using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading before the list.
        builder.Writeln("Features of Aspose.Words:");

        // Start a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add list items – each call to Writeln creates a new paragraph that becomes a list item.
        builder.Writeln("High performance");
        builder.Writeln("Robust API");
        builder.Writeln("Extensive documentation");
        builder.Writeln("Cross‑platform support");
        builder.Writeln("Easy integration");

        // End the bulleted list – removes bullets and resets the list level.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph after the list.
        builder.Writeln("End of feature list.");

        // Save the document to a DOCX file.
        doc.Save("ListExample.docx");
    }
}
