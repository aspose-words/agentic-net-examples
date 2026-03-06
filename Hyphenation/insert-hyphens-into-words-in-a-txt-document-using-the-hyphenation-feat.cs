using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateTxtDocument
{
    static void Main()
    {
        // Path to the folder containing the input TXT file.
        string dataDir = @"C:\Data\";

        // Load the TXT document.
        string inputPath = Path.Combine(dataDir, "input.txt");
        Document doc = new Document(inputPath);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: adjust hyphenation zone (distance from right margin where hyphenation is suppressed).
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 20 points)

        // Save the hyphenated document. Saving to DOCX preserves layout and inserted hyphens.
        string outputPath = Path.Combine(dataDir, "output.docx");
        doc.Save(outputPath);
    }
}
