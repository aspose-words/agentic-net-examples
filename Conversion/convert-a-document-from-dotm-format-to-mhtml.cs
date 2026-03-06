using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDotmToMhtml
{
    static void Main()
    {
        // Path to the source DOTM file.
        string inputPath = "input.dotm";

        // Path where the MHTML file will be saved.
        string outputPath = "output.mht";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Option 1: Directly save using the SaveFormat overload.
        doc.Save(outputPath, SaveFormat.Mhtml);

        // Option 2: Use HtmlSaveOptions for more control (uncomment if needed).
        // HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        // doc.Save(outputPath, saveOptions);
    }
}
