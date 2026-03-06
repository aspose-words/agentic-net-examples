using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace StyleSeparatorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply a built‑in heading style to the first part of the line.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("This text is in a Heading style. ");

            // Insert a style separator. This creates a new paragraph internally,
            // but there will be no visible line break between the two parts.
            builder.InsertStyleSeparator();

            // Create a custom paragraph style.
            Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 8;
            customStyle.Font.Bold = false;

            // Apply the custom style to the second part of the line.
            builder.ParagraphFormat.StyleName = customStyle.Name;
            builder.Write("This text is in a custom style.");

            // Save the document to a DOCX file.
            doc.Save("StyleSeparator.docx");
        }
    }
}
