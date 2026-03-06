using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph containing long words that will be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("Antidisestablishmentarianism is a long word that often needs hyphenation when it reaches the end of a line.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: define the hyphenation zone (distance from the right margin) and the maximum
        // number of consecutive lines that may end with a hyphen.
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 1440 points)
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Save the document as a DOT (Word template) file.
        doc.Save("HyphenatedTemplate.dot");
    }
}
