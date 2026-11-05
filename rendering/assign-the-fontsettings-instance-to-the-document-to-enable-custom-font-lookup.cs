using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the output document and a custom fonts folder.
        string outputPath = "RenderedDocument.pdf";
        string fontsFolder = "CustomFonts";

        // Ensure the custom fonts folder exists (it can be empty for this demo).
        Directory.CreateDirectory(fontsFolder);

        // Create a FontSettings instance and point it to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Create a new document and add a paragraph that uses a font name that may require the custom folder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFont"; // Example font name.
        builder.Writeln("This paragraph is rendered using a custom font lookup.");

        // Assign the FontSettings instance to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The rendered PDF was not created.", outputPath);

        // Indicate successful completion.
        Console.WriteLine("Document rendered and saved to: " + Path.GetFullPath(outputPath));
    }
}
