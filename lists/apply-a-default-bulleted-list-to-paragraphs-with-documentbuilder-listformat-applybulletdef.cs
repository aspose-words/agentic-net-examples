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

        // Write a heading before the list.
        builder.Writeln("Advantages of Aspose.Words:");

        // Start a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add list items.
        builder.Writeln("Great performance");
        builder.Writeln("High reliability");
        builder.Writeln("Rich feature set");
        builder.Writeln("Easy to use API");
        builder.Writeln("Extensive documentation");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ApplyDefaultBullets.docx");
        doc.Save(outputPath);
    }
}
