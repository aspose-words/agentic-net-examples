using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class HighlightParagraphExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a normal introductory line.
        builder.Writeln("Important information:");

        // Insert a new paragraph where the highlighted text will go.
        builder.InsertParagraph();

        // Apply solid shading with a light yellow background to the current paragraph.
        builder.ParagraphFormat.Shading.Texture = TextureIndex.TextureSolid;
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Write the highlighted content.
        builder.Writeln("This paragraph is highlighted with a light yellow background.");

        // (Optional) Reset shading for any following paragraphs.
        builder.ParagraphFormat.Shading.ClearFormatting();

        // Save the document to a file.
        doc.Save("HighlightedParagraph.docx");
    }
}
