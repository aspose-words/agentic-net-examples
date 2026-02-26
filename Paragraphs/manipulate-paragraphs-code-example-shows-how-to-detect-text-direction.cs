using System;
using Aspose.Words;
using Aspose.Words.Loading;

class DetectTextDirection
{
    static void Main()
    {
        // Path to the folder containing the input TXT files.
        string inputFolder = @"C:\InputTxtFiles\";

        // Create a TxtLoadOptions object to configure how the TXT file is loaded.
        TxtLoadOptions loadOptions = new TxtLoadOptions();

        // Enable automatic detection of paragraph direction.
        loadOptions.DocumentDirection = DocumentDirection.Auto;

        // Load a Hebrew text file (right‑to‑left).
        Document hebrewDoc = new Document(System.IO.Path.Combine(inputFolder, "Hebrew text.txt"), loadOptions);
        bool hebrewIsRtl = hebrewDoc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi;
        Console.WriteLine($"Hebrew paragraph is RTL: {hebrewIsRtl}");

        // Load an English text file (left‑to‑right).
        Document englishDoc = new Document(System.IO.Path.Combine(inputFolder, "English text.txt"), loadOptions);
        bool englishIsRtl = englishDoc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi;
        Console.WriteLine($"English paragraph is RTL: {englishIsRtl}");

        // Optionally, save the loaded documents as DOCX files.
        string outputFolder = @"C:\OutputDocs\";
        hebrewDoc.Save(System.IO.Path.Combine(outputFolder, "Hebrew.docx"));
        englishDoc.Save(System.IO.Path.Combine(outputFolder, "English.docx"));
    }
}
