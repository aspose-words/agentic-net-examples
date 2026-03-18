using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Enable automatic hyphenation for the entire document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings if desired.
        // doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // doc.HyphenationOptions.HyphenationZone = 720;
        // doc.HyphenationOptions.HyphenateCaps = true;

        // Define the output file path. Adjust the folder as needed.
        string outputPath = "HyphenatedDocument.docx";

        // Save the document in DOCX format. The file extension determines the format.
        doc.Save(outputPath);
    }
}
