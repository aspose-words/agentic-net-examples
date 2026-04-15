using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 0.25 inches = 18 points. Set a hanging indent by using a negative FirstLineIndent.
        // Also set LeftIndent to keep the subsequent lines aligned with the hanging indent.
        builder.ParagraphFormat.FirstLineIndent = -18; // negative value creates hanging indent
        builder.ParagraphFormat.LeftIndent = 18;      // indent the whole paragraph

        // Add a sample citation paragraph.
        builder.Writeln("Doe, J. (2023). Example citation for a reference list.");

        // Save the document to the local file system.
        string outputPath = "HangingIndent.docx";
        doc.Save(outputPath);
    }
}
