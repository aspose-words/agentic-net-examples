using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string inputFile = @"C:\Docs\SourceDocument.docx";

        // Path where the PNG image will be saved.
        const string outputFile = @"C:\Docs\SourceDocument.png";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // ----- FAQ processing (example: replace a placeholder with an answer) -----
        // This demonstrates a simple find‑and‑replace operation.
        // Adjust the search pattern and replacement text as needed for your FAQ logic.
        const string placeholder = "?question?";
        const string answer = "This is the answer to the frequently asked question.";
        doc.Range.Replace(placeholder, answer, new FindReplaceOptions());

        // Configure image save options to render the first page as a PNG.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page (zero‑based index).
            PageSet = new PageSet(0),

            // Optional: increase resolution for higher quality output.
            Resolution = 300
        };

        // Save the document as a PNG image.
        doc.Save(outputFile, saveOptions);
    }
}
