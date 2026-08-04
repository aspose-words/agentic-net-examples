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

        // Set the drop cap height to span 4 lines.
        builder.ParagraphFormat.LinesToDrop = 4;
        builder.Writeln("H"); // This paragraph becomes the drop cap.

        // Reset the drop cap setting for subsequent paragraphs.
        builder.ParagraphFormat.LinesToDrop = 0;
        builder.Writeln("ello world!"); // Normal paragraph that wraps around the drop cap.

        // Save the document to the output folder.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DropCapExample.docx");
        doc.Save(outputPath);
    }
}
