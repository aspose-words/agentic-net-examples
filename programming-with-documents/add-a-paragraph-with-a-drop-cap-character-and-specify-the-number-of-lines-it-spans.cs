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

        // Configure the first paragraph as a drop cap.
        // The character will occupy the height of 4 lines.
        builder.Font.Size = 48; // make the drop cap large
        builder.ParagraphFormat.LinesToDrop = 4;
        builder.ParagraphFormat.DropCapPosition = DropCapPosition.Normal; // optional positioning
        builder.Writeln("L"); // drop cap character

        // Reset the paragraph formatting for the following text.
        builder.ParagraphFormat.LinesToDrop = 0;
        builder.Font.Size = 12;
        builder.Writeln("orem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphDropCap.docx");
        doc.Save(outputPath);
    }
}
