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
        // Increase font size for the drop cap character.
        builder.Font.Size = 48;
        builder.Writeln("D");

        // Return to normal paragraph formatting.
        builder.ParagraphFormat.LinesToDrop = 0;
        builder.Font.Size = 12;
        builder.Writeln("rop cap example text that wraps around the large initial letter.");

        // Prepare output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DropCapExample.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
