using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class StyleIdentifierExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve a built‑in style using its locale‑independent identifier.
        // This works regardless of the language of the document.
        Style headingStyle = doc.Styles[StyleIdentifier.Heading1];

        // Example modification: change the font size of the Heading 1 style.
        headingStyle.Font.Size = 16;

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
