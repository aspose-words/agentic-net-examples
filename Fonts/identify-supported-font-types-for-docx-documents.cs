using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Get the collection of fonts defined in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        Console.WriteLine($"The document defines {fontInfos.Count} font(s):");

        // Iterate through each FontInfo and display its characteristics.
        for (int i = 0; i < fontInfos.Count; i++)
        {
            FontInfo font = fontInfos[i];

            // Determine whether the font is a TrueType/OpenType font.
            string fontType = font.IsTrueType ? "TrueType/OpenType" : "Raster/Vector";

            Console.WriteLine($"- Name   : {font.Name}");
            Console.WriteLine($"  Family : {font.Family}");
            Console.WriteLine($"  Type   : {fontType}");
            Console.WriteLine($"  Pitch  : {font.Pitch}");
        }
    }
}
