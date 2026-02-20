using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load a DOCX document.
        Document doc = new Document("Input.docx");

        // Get the collection of fonts used in the document.
        FontInfoCollection fonts = doc.FontInfos;

        // Iterate through each FontInfo and display its supported type information.
        foreach (FontInfo font in fonts)
        {
            // Font name.
            Console.WriteLine($"Font name: {font.Name}");

            // TrueType/OpenType flag – true if the font is a TrueType or OpenType font.
            Console.WriteLine($"Is TrueType/OpenType: {font.IsTrueType}");

            // Font family (Roman, Swiss, Modern, etc.).
            Console.WriteLine($"Family: {font.Family}");

            // Font pitch (Fixed, Variable, Default).
            Console.WriteLine($"Pitch: {font.Pitch}");

            // Character set (e.g., 0 = ANSI).
            Console.WriteLine($"Charset: {font.Charset}");

            Console.WriteLine(new string('-', 40));
        }
    }
}
