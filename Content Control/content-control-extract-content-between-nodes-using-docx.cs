using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("Input.docx");

        // Locate the first content control (StructuredDocumentTag) in the document.
        // The GetChild method searches the whole document tree when the third argument is true.
        StructuredDocumentTag contentControl = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

        // If a content control was found, extract its inner text via the Range property.
        string extractedText = contentControl?.Range?.Text ?? string.Empty;

        // Output the extracted text to the console.
        Console.WriteLine(extractedText);
    }
}
