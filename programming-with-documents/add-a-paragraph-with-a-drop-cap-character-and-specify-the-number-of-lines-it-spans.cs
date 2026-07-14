using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the drop cap to span 4 lines.
        builder.ParagraphFormat.LinesToDrop = 4;
        // Insert the drop cap character.
        builder.Writeln("H");

        // Reset the drop cap setting for normal text.
        builder.ParagraphFormat.LinesToDrop = 0;
        // Add the remaining paragraph text.
        builder.Writeln("ello world! This paragraph wraps around the drop cap character.");

        // Prepare an output folder and file name.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);
        string outputFile = Path.Combine(outputFolder, "DropCapExample.docx");

        // Save the document.
        doc.Save(outputFile);
    }
}
