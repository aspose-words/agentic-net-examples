using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OpenTypePreserveExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing Word document.
            Document doc = new Document("Input.docx");

            // Ensure OpenType font formatting features (ligatures, contextual forms, etc.) are NOT disabled.
            // Setting this property to false preserves the typographic features during PDF conversion.
            doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = false;

            // Save the document as PDF, preserving OpenType features.
            doc.Save("Output.pdf", SaveFormat.Pdf);
        }
    }
}
