using System;
using System.IO;
using Aspose.Words;

namespace InsertHyphensIntoTxt
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the input file.
            string dataDir = @"C:\Data\";

            // Load a plain‑text file into an Aspose.Words Document.
            Document doc = new Document(Path.Combine(dataDir, "input.txt"));

            // Enable automatic hyphenation so Aspose.Words inserts discretionary hyphens where needed.
            doc.HyphenationOptions.AutoHyphenation = true;
            // Optional: adjust hyphenation settings (zone, consecutive limit, etc.) as required.
            doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // limit consecutive hyphenated lines.
            doc.HyphenationOptions.HyphenateCaps = true;         // hyphenate all‑caps words.

            // When saving to plain text the discretionary hyphen (U+00AD) is not visible.
            // Replace it with a regular hyphen so the output file shows the inserted hyphens.
            doc.Range.Replace("\u00AD", "-");

            // Save the document back to TXT; the inserted hyphens are now visible.
            doc.Save(Path.Combine(dataDir, "output.txt"));
        }
    }
}
