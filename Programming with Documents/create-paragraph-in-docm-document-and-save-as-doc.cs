using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document (DOCM format will be set when saving as DOCM if needed)
            Document doc = new Document();

            // Use DocumentBuilder to add a paragraph to the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a new paragraph added to the document.");

            // Save the document as a DOC (Word 97-2007) file.
            // The Save method automatically determines the format from the extension,
            // but we explicitly specify SaveFormat.Doc for clarity.
            doc.Save("Output.doc", SaveFormat.Doc);
        }
    }
}
