using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

namespace HyphenationPdfComparison
{
    class Program
    {
        static void Main()
        {
            // Prepare sample text that will benefit from hyphenation.
            const string sampleText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                                      "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";

            // ---------- Document with automatic hyphenation ----------
            Document hyphenatedDoc = new Document();                     // create document
            DocumentBuilder hyphenatedBuilder = new DocumentBuilder(hyphenatedDoc);
            hyphenatedBuilder.Font.Size = 24;                           // make text larger to force line breaks
            hyphenatedBuilder.Writeln(sampleText);

            // Enable automatic hyphenation.
            hyphenatedDoc.HyphenationOptions.AutoHyphenation = true;
            hyphenatedDoc.HyphenationOptions.HyphenationZone = 720;    // optional: increase zone
            hyphenatedDoc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            hyphenatedDoc.HyphenationOptions.HyphenateCaps = true;

            // Save as PDF.
            string hyphenatedPdfPath = Path.Combine(Path.GetTempPath(), "Hyphenated.pdf");
            hyphenatedDoc.Save(hyphenatedPdfPath);                     // save document (PDF inferred from extension)

            // ---------- Document without hyphenation ----------
            Document plainDoc = new Document();                          // create second document
            DocumentBuilder plainBuilder = new DocumentBuilder(plainDoc);
            plainBuilder.Font.Size = 24;
            plainBuilder.Writeln(sampleText);

            // Ensure hyphenation is disabled (default is false, but set explicitly for clarity).
            plainDoc.HyphenationOptions.AutoHyphenation = false;

            // Save as PDF.
            string plainPdfPath = Path.Combine(Path.GetTempPath(), "Plain.pdf");
            plainDoc.Save(plainPdfPath);

            // ---------- Compare file sizes ----------
            long hyphenatedSize = new FileInfo(hyphenatedPdfPath).Length;
            long plainSize = new FileInfo(plainPdfPath).Length;

            Console.WriteLine($"Hyphenated PDF size : {hyphenatedSize} bytes");
            Console.WriteLine($"Non‑hyphenated PDF size : {plainSize} bytes");

            if (hyphenatedSize < plainSize)
                Console.WriteLine("Hyphenation reduced the PDF file size.");
            else if (hyphenatedSize > plainSize)
                Console.WriteLine("Hyphenation increased the PDF file size.");
            else
                Console.WriteLine("Both PDFs have the same size.");
        }
    }
}
