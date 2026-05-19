using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write an introductory paragraph.
        builder.Writeln("Aspose.Words main advantages are:");

        // Start a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add bullet items.
        builder.Writeln("High performance");
        builder.Writeln("Rich API");
        builder.Writeln("Cross‑platform support");
        builder.Writeln("Extensive documentation");
        builder.Writeln("Easy to use");

        // End the bulleted list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "BulletedList.docx");
        doc.Save(outputPath);
    }
}
