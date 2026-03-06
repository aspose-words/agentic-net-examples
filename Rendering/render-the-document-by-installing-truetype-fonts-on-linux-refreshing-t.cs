using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class RenderDocumentWithLinuxFonts
{
    static void Main()
    {
        // ---------------------------------------------------------------------
        // 1. Path to the folder that contains the TrueType fonts on the Linux machine.
        //    The fonts must be installed or copied to this directory beforehand.
        // ---------------------------------------------------------------------
        string fontsFolder = "/usr/share/fonts/truetype/custom";

        // Verify that the folder exists.
        if (!Directory.Exists(fontsFolder))
        {
            Console.WriteLine($"Fonts folder not found: {fontsFolder}");
            return;
        }

        // ---------------------------------------------------------------------
        // 2. Refresh the Linux font cache so the system recognises the new fonts.
        //    This step is performed outside of the .NET code, e.g.:
        //        $ sudo fc-cache -f -v
        //    The code below only tells Aspose.Words where to look for the fonts.
        // ---------------------------------------------------------------------

        // 3. Configure Aspose.Words to use the custom fonts folder.
        //    The second argument (true) tells Aspose.Words to search the folder recursively.
        FontSettings.DefaultInstance.SetFontsFolder(fontsFolder, recursive: true);

        // ---------------------------------------------------------------------
        // 4. Load the source document. No LoadOptions are required because the
        //    default FontSettings instance has already been configured.
        // ---------------------------------------------------------------------
        string sourceDocPath = "input.docx";
        if (!File.Exists(sourceDocPath))
        {
            Console.WriteLine($"Source document not found: {sourceDocPath}");
            return;
        }

        Document doc = new Document(sourceDocPath);

        // ---------------------------------------------------------------------
        // 5. Render the document to PDF (or any other fixed‑page format).
        // ---------------------------------------------------------------------
        string outputPdfPath = "output.pdf";
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        Console.WriteLine($"Document rendered successfully to: {outputPdfPath}");
    }
}
