using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to examine.
        string docPath = @"MyDir\Document.docx";

        // Load the DOCX document.
        Document doc = new Document(docPath);

        // Retrieve the collection of fonts defined in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        Console.WriteLine($"Document contains {fontInfos.Count} font definition(s).");

        // Iterate through each FontInfo and display its key properties.
        for (int i = 0; i < fontInfos.Count; i++)
        {
            FontInfo font = fontInfos[i];
            Console.WriteLine($"Font #{i + 1}");
            Console.WriteLine($"  Name   : {font.Name}");
            Console.WriteLine($"  Family : {font.Family}");          // FontFamily enum
            Console.WriteLine($"  TrueType? : {font.IsTrueType}");
            Console.WriteLine($"  Pitch  : {font.Pitch}");           // FontPitch enum
            Console.WriteLine($"  Charset: {font.Charset}");
        }

        // List the supported FontFamily enumeration values for DOCX.
        Console.WriteLine("\nSupported FontFamily values:");
        foreach (FontFamily family in Enum.GetValues(typeof(FontFamily)))
        {
            Console.WriteLine($"  {family} = {(int)family}");
        }

        // List the supported FontPitch enumeration values for DOCX.
        Console.WriteLine("\nSupported FontPitch values:");
        foreach (FontPitch pitch in Enum.GetValues(typeof(FontPitch)))
        {
            Console.WriteLine($"  {pitch} = {(int)pitch}");
        }
    }
}
