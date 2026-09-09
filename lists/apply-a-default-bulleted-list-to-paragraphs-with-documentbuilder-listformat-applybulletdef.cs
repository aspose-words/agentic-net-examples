using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading before the list.
        builder.Writeln("Aspose.Words main advantages are:");

        // Start a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add list items that will be formatted with bullets.
        builder.Writeln("High performance");
        builder.Writeln("Robust API");
        builder.Writeln("Cross‑platform support");
        builder.Writeln("Extensive format conversion");
        builder.Writeln("Easy to use");

        // End the bulleted list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DefaultBulletedList.docx");
        doc.Save(outputPath);
    }
}
