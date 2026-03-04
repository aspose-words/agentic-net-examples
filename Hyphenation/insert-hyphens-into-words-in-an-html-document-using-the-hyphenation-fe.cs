using System;
using Aspose.Words;

namespace HyphenateHtmlExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the input HTML file.
            string dataDir = @"C:\MyData\";

            // Load the existing HTML document.
            Document doc = new Document(dataDir + "input.html");

            // Enable automatic hyphenation for the document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Optional: fine‑tune hyphenation behavior.
            doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
            doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate words in all caps.

            // Save the document back to HTML.
            // Aspose.Words inserts soft‑hyphen characters (U+00AD) where hyphenation occurs.
            doc.Save(dataDir + "output.html");
        }
    }
}
