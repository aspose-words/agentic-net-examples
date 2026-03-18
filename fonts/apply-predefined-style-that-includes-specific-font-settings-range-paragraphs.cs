using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Replacing;

class ApplyStyleToParagraphRange
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a custom paragraph style named "MyCustomStyle".
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        // Define the font settings for the style.
        customStyle.Font.Name = "Calibri";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = Color.DarkBlue;
        customStyle.Font.Bold = true;
        customStyle.Font.Italic = true;

        // Use DocumentBuilder to add several paragraphs to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraphs that will keep the default style.
        builder.Writeln("Paragraph 1 – default style.");
        builder.Writeln("Paragraph 2 – default style.");

        // Switch to the custom style for the next three paragraphs (the target range).
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("Paragraph 3 – custom style applied.");
        builder.Writeln("Paragraph 4 – custom style applied.");
        builder.Writeln("Paragraph 5 – custom style applied.");

        // Return to the default style for subsequent paragraphs.
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("Paragraph 6 – default style again.");

        // Alternatively, if you need to apply the style to an existing range of paragraphs
        // after the document has been built, you can iterate and set the style manually.
        // Here we apply the custom style to paragraphs 2 through 4 (zero‑based index).
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
        for (int i = 1; i <= 3 && i < paragraphs.Count; i++)
        {
            paragraphs[i].ParagraphFormat.Style = customStyle;
        }

        // Save the document to disk.
        doc.Save("StyledParagraphs.docx");
    }
}
