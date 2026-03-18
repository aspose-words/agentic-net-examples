using System;
using Aspose.Words;

namespace HyphenationExample
{
    class Program
    {
        static void Main()
        {
            // Path where the document will be saved.
            string outputPath = "HyphenatedDocument.docx";

            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some text that can be hyphenated.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Enable automatic hyphenation and configure optional settings.
            doc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;     // Max consecutive hyphenated lines.
            doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch zone from right margin.
            doc.HyphenationOptions.HyphenateCaps = true;           // Hyphenate all‑caps words.

            // Save the document to DOCX format.
            doc.Save(outputPath);

            // Reload the saved document to verify that hyphenation settings persisted.
            Document loadedDoc = new Document(outputPath);

            // Simple verification: the AutoHyphenation flag should still be true.
            if (!loadedDoc.HyphenationOptions.AutoHyphenation)
                throw new InvalidOperationException("Hyphenation settings were not retained after saving.");

            // Additional verification could involve layout inspection, but checking the flag suffices.
            Console.WriteLine("Hyphenation enabled and retained successfully.");
        }
    }
}
