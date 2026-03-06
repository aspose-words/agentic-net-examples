using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the PNG image will be saved.
        const string outputPath = @"C:\Docs\ResultImage.png";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Ensure the document contains at least one paragraph.
        doc.EnsureMinimum();
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Clear any existing runs and insert new text using a Run node.
        firstParagraph.Runs.Clear();
        firstParagraph.AppendChild(new Run(doc, "This text was inserted using the Text property."));

        // Save the modified document as a PNG image (renders the first page).
        doc.Save(outputPath, SaveFormat.Png);
    }
}
