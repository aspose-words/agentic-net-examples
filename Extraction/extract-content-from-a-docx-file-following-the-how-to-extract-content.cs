using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.AI; // Only needed if AI features are used; not required here.

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\Sample.docx";

        // Load the Word document using the Document constructor (load rule).
        Document doc = new Document(sourcePath);

        // Extract plain‑text representation of the document.
        // PlainTextDocument constructor loads the file and provides the Text property.
        PlainTextDocument plainText = new PlainTextDocument(sourcePath);
        string extractedText = plainText.Text;

        // Output the extracted text to the console.
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);

        // Optionally, save the extracted text to a .txt file.
        string outputPath = @"C:\Docs\SampleExtracted.txt";
        File.WriteAllText(outputPath, extractedText);
    }
}
