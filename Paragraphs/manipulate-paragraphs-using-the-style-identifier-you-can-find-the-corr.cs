using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve a built‑in style using its locale‑independent identifier.
        // This works regardless of the language of the document.
        Style headingStyle = doc.Styles[StyleIdentifier.Heading1];

        // Example modification: change the font size of the Heading1 style.
        headingStyle.Font.Size = 16;

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // If the paragraph currently uses the "Normal" style, change it to Heading1.
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Normal)
            {
                // Apply the style by setting the identifier (language‑independent).
                para.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
