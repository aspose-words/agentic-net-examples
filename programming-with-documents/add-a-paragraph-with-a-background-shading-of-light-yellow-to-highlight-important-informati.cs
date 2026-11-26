using System;
using System.IO;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading to light yellow.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        // No need to set texture; default is None.

        // Add the highlighted paragraph.
        builder.Writeln("Important information");

        // Define output path.
        string outputPath = "HighlightedParagraph.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created and can be reopened.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // Document loaded successfully; program ends here.
        }
    }
}
