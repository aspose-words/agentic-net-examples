using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Access the collection of fonts defined in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        Console.WriteLine($"Document contains {fontInfos.Count} font definition(s).");

        // Iterate through each FontInfo and display its supported properties.
        for (int i = 0; i < fontInfos.Count; i++)
        {
            FontInfo font = fontInfos[i];

            Console.WriteLine($"Font #{i + 1}:");
            Console.WriteLine($"  Name       : {font.Name}");
            Console.WriteLine($"  IsTrueType : {font.IsTrueType}");
            Console.WriteLine($"  Family     : {font.Family}");
            Console.WriteLine($"  Pitch      : {font.Pitch}");
            Console.WriteLine($"  Charset    : {font.Charset}");
        }
    }
}
