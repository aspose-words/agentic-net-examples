using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        string inputPath = "Input.docx";
        Document doc = new Document(inputPath);

        // Delete all existing content in the document's main range.
        doc.Range.Delete();

        // Insert the new text into the now‑empty document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the new content that replaces the original document.");

        // Save the updated document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
