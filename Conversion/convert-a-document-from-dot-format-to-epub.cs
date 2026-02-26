using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document in DOT (Word template) format.
        Document doc = new Document("input.dot");

        // Save the document as EPUB using the SaveFormat enumeration.
        doc.Save("output.epub", SaveFormat.Epub);
    }
}
