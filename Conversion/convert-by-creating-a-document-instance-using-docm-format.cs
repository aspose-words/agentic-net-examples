using System;
using Aspose.Words;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Optional: add some content using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, DOCM world!");

            // Save the document in the macro‑enabled DOCM format.
            // The SaveFormat.Docm enum value specifies the DOCM file type.
            doc.Save("Result.docm", SaveFormat.Docm);
        }
    }
}
