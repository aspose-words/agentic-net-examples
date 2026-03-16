using System;
using Aspose.Words;

namespace HyphenationExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Path where the hyphenated DOCX will be saved.
            string outputPath = @"C:\Docs\OutputDocument_Hyphenated.docx";

            // Load the existing document from disk.
            Document doc = new Document(inputPath);

            // Enable automatic hyphenation for the whole document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Optional: configure additional hyphenation settings.
            // Maximum number of consecutive lines that may end with hyphens.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

            // Distance from the right margin within which hyphenation is suppressed (0.25 inch = 360 twips).
            doc.HyphenationOptions.HyphenationZone = 360;

            // Hyphenate words written in all capital letters.
            doc.HyphenationOptions.HyphenateCaps = true;

            // Save the document back to DOCX format.
            // The hyphenation marks generated during layout are preserved in the saved file.
            doc.Save(outputPath);
        }
    }
}
