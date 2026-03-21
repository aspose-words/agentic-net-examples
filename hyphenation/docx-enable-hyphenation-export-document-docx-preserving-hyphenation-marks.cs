using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HyphenationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document with some sample text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample paragraph containing some longwordswhichmightneedhyphenation to demonstrate automatic hyphenation in Aspose.Words.");
            builder.Writeln("Another line with CAPITALIZEDWORDS that should also be hyphenated when the option is enabled.");

            // Enable automatic hyphenation for the whole document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Optional: configure additional hyphenation settings.
            // Maximum number of consecutive lines that may end with hyphens.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

            // Distance from the right margin within which hyphenation is suppressed (0.25 inch = 360 twips).
            doc.HyphenationOptions.HyphenationZone = 360;

            // Hyphenate words written in all capital letters.
            doc.HyphenationOptions.HyphenateCaps = true;

            // Determine output path in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocument_Hyphenated.docx");

            // Save the document. The hyphenation marks generated during layout are preserved in the saved file.
            doc.Save(outputPath);

            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
