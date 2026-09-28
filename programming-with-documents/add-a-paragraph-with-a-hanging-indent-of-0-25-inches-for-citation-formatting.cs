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

        // Set hanging indent: left indent 0.25 inches (18 points), first line indent -0.25 inches.
        builder.ParagraphFormat.LeftIndent = 18;          // 0.25 inches = 18 points
        builder.ParagraphFormat.FirstLineIndent = -18;   // negative for hanging indent

        // Add the citation paragraph.
        builder.Writeln("This is a citation that requires a hanging indent.");

        // Save the document.
        string outputPath = "HangingIndent.docx";
        doc.Save(outputPath);

        // Confirm the file was saved.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
