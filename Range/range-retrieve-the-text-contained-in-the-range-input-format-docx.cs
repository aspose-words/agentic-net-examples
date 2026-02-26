using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file whose text we want to extract.
        string docPath = "input.docx";

        // Load the document using the provided Document(string) constructor.
        Document doc = new Document(docPath);

        // Retrieve the complete text of the document's range.
        string text = doc.Range.Text;

        // Display the extracted text.
        Console.WriteLine(text);
    }
}
