using System;
using Aspose.Words;

namespace AsposeWordsDotxExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Optionally add some content using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello from a DOTX template!");

            // Save the document as a DOTX (Word template) file.
            // The Save method with (string, SaveFormat) follows the provided rule.
            doc.Save("OutputTemplate.dotx", SaveFormat.Dotx);

            // Load the saved DOTX file back into a Document object.
            // This uses the Document(string) constructor rule.
            Document loadedDoc = new Document("OutputTemplate.dotx");

            // Verify that the content was loaded correctly (optional).
            Console.WriteLine(loadedDoc.GetText().Trim());
        }
    }
}
