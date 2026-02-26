using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOTM file.
        string inputPath = "input.dotm";

        // Path where the MHTML file will be saved.
        string outputPath = "output.mht";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Save the document in MHTML format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
