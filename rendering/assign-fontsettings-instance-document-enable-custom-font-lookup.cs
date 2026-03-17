using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new empty document
        Document doc = new Document();

        // Configure custom font lookup
        FontSettings fontSettings = new FontSettings();
        // (Optional) Add a custom font folder if you have fonts in a specific directory
        // fontSettings.SetFontsFolder(@"C:\MyCustomFonts", recursive: true);

        // Assign the FontSettings to the document
        doc.FontSettings = fontSettings;

        // Save the document to PDF (uses the assigned FontSettings)
        doc.Save("Output.pdf");
    }
}
