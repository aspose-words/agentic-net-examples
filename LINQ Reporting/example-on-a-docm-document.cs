using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsDocmExample
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new blank document (DOCM - macro-enabled format)
            // -----------------------------------------------------------------
            Document doc = new Document(); // uses the Document() constructor

            // Use DocumentBuilder to add some content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, this is a macro‑enabled DOCM document created with Aspose.Words.");

            // Insert a simple shape to demonstrate that the document can contain
            // regular Word objects as well as macros (if added later).
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // -----------------------------------------------------------------
            // 2. Save the document as a DOCM file.
            // -----------------------------------------------------------------
            // The Save(string) method automatically determines the format from the file extension.
            string outputPath = "OutputDocument.docm";
            doc.Save(outputPath); // saves as DOCM because of the .docm extension

            // -----------------------------------------------------------------
            // 3. Load the previously saved DOCM document.
            // -----------------------------------------------------------------
            // Demonstrates loading a macro‑enabled document using the Document(string) constructor.
            Document loadedDoc = new Document(outputPath);

            // Verify that the loaded document contains the expected text.
            string text = loadedDoc.GetText().Trim();
            Console.WriteLine("Loaded document text:");
            Console.WriteLine(text);
        }
    }
}
