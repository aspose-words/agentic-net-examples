using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;

class ListToPlainText
{
    static void Main()
    {
        // Path to the source DOC/DOCX file that contains the list.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Path where the plain‑text file will be saved.
        string targetPath = @"C:\Docs\ListAsText.txt";

        // Load the Word document.
        Document doc = new Document(sourcePath);

        // Ensure that list labels are up‑to‑date before exporting.
        doc.UpdateListLabels();

        // Configure text‑save options.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Simplify complex list symbols to plain ASCII characters.
            SimplifyListLabels = true,

            // Optional: define how list indentation is represented.
            // Here we use three spaces per level.
            ListIndentation = { Character = ' ', Count = 3 }
        };

        // Save the document as plain text using the configured options.
        doc.Save(targetPath, txtOptions);
    }
}
