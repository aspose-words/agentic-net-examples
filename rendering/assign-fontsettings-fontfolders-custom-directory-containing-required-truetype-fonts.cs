using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the required TrueType fonts.
        string customFontsFolder = @"C:\MyFonts";

        // Create a new blank document. (Replace with loading an existing document if needed.)
        Document doc = new Document();

        // Configure the static default FontSettings instance to look for fonts in the custom folder.
        // The second argument 'true' enables recursive scanning of subfolders.
        FontSettings.DefaultInstance.SetFontsFolder(customFontsFolder, true);

        // Assign the configured FontSettings to the document (optional when using DefaultInstance).
        doc.FontSettings = FontSettings.DefaultInstance;

        // Render the document to a PDF file (or any other supported format).
        doc.Save("Output.pdf");
    }
}
