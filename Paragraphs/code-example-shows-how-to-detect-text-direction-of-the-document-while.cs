using System;
using Aspose.Words;
using Aspose.Words.Loading;

class DetectTextDirection
{
    static void Main()
    {
        // Folder that contains the source TXT files.
        string dataDir = @"C:\Data\";

        // Create load options and enable automatic detection of text direction.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.DocumentDirection = DocumentDirection.Auto;

        // Load a Hebrew text file. The paragraph will be marked as right‑to‑left.
        Document hebrewDoc = new Document(dataDir + "Hebrew.txt", loadOptions);
        bool hebrewIsRtl = hebrewDoc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi;
        Console.WriteLine($"Hebrew paragraph Bidi: {hebrewIsRtl}");

        // Load an English text file. The paragraph will be left‑to‑right.
        Document englishDoc = new Document(dataDir + "English.txt", loadOptions);
        bool englishIsRtl = englishDoc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi;
        Console.WriteLine($"English paragraph Bidi: {englishIsRtl}");

        // Save the resulting DOCX files.
        hebrewDoc.Save(dataDir + "Hebrew.docx");
        englishDoc.Save(dataDir + "English.docx");
    }
}
