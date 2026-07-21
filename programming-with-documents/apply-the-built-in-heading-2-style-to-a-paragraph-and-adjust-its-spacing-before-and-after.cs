using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading 2 style to the next paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Adjust spacing before and after the paragraph (values are in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 6;   // 6 points after

        // Write the paragraph text.
        builder.Writeln("This is a Heading 2 styled paragraph with custom spacing.");

        // Define an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Heading2_Styled.docx");

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
