using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document from the file system.
            // The Document constructor handles loading and format detection.
            Document doc = new Document("input.docx");

            // Create a DocumentBuilder tied to the loaded document.
            // This allows us to modify the document's content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text at the current cursor position using the Write method.
            // The Write method adds the specified string directly into the document.
            builder.Write("Inserted text using the Text property.");

            // Save the modified document as a PNG image.
            // The Save method determines the format from the SaveFormat enum.
            doc.Save("output.png", SaveFormat.Png);
        }
    }
}
