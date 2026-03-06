using System;
using System.IO;
using Aspose.Words;

class ExtractRangeText
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "input.docx";

        // Load the document using the provided Document(string) constructor (load rule).
        Document doc = new Document(sourcePath);

        // Extract the text of the whole document range.
        // The Range.Text property returns the concatenated text of all nodes in the range.
        string extractedText = doc.Range.Text;

        // Output the extracted text to the console.
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);

        // Optionally, save the original document (demonstrating the provided Save(string) rule).
        string savedDocPath = "saved_copy.docx";
        doc.Save(savedDocPath);

        // Optionally, write the extracted text to a plain text file.
        // This uses standard .NET I/O and does not replace the required lifecycle rules.
        string textFilePath = "extracted_text.txt";
        File.WriteAllText(textFilePath, extractedText);
    }
}
