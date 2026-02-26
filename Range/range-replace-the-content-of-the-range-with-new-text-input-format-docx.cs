using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file whose content will be replaced.
        const string inputPath = "input.docx";

        // Path where the modified document will be saved.
        const string outputPath = "output.docx";

        // The new text that will replace the entire content of the document.
        const string newText = "Hello Aspose! This is the new document content.";

        // Load the existing document (lifecycle rule: load).
        Document doc = new Document(inputPath);

        // Remove all existing characters from the document's main range (lifecycle rule: use Range.Delete).
        doc.Range.Delete();

        // Insert the new text into the now‑empty document using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write(newText);

        // Save the modified document (lifecycle rule: save).
        doc.Save(outputPath);
    }
}
