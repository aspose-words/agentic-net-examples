using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainText
{
    static void Main()
    {
        // Path to the source DOC document that contains the list.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting plain‑text file will be saved.
        string outputPath = @"C:\Docs\ListPlainText.txt";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Ensure that list labels are up‑to‑date before exporting.
        doc.UpdateListLabels();

        // Configure text‑save options.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Simplify list symbols to basic ASCII characters (e.g., "*", "1.", etc.).
            SimplifyListLabels = true,

            // Optional: define how list indentation is represented.
            // Here we use three spaces per level.
            ListIndentation = { Character = ' ', Count = 3 }
        };

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, txtOptions);
    }
}
