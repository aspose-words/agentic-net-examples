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

        // Write a paragraph. The first character will become the drop cap.
        builder.Writeln("A quick brown fox jumps over the lazy dog.");

        // Retrieve the paragraph we just added.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Enable drop cap. The default number of lines is 3, which satisfies the requirement.
        paragraph.ParagraphFormat.DropCapPosition = DropCapPosition.Normal;

        // Save the document to disk.
        string outputPath = "DropCapExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
