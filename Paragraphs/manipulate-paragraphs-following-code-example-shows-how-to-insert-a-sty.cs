using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading1 style to the first part of the line.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator. This creates a hidden paragraph break that
        // allows the next text to have a different paragraph style while staying on the same line.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Bold = false;
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";

        // Apply the custom style to the second part of the line.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Verify that two paragraphs were created and that their styles are correct.
        if (doc.FirstSection.Body.Paragraphs.Count != 2 ||
            doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Style.Name != "Heading 1" ||
            doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.Style.Name != "MyParaStyle")
        {
            throw new InvalidOperationException("Style separator did not work as expected.");
        }

        // Save the document.
        doc.Save("StyleSeparatorExample.docx");
    }
}
