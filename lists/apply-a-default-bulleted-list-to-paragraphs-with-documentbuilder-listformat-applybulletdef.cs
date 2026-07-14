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

        // Initialize DocumentBuilder which will be used to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write an introductory paragraph.
        builder.Writeln("Features of Aspose.Words:");

        // Start a default bulleted list and apply it to subsequent paragraphs.
        builder.ListFormat.ApplyBulletDefault();

        // Add list items.
        builder.Writeln("High performance");
        builder.Writeln("Robust API");
        builder.Writeln("Extensive format support");
        builder.Writeln("Easy to integrate");
        builder.Writeln("Comprehensive documentation");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BulletedList.docx");
        doc.Save(outputPath);
    }
}
