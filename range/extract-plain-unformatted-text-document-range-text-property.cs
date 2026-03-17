using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document from the file system.
        Document doc = new Document("input.docx");

        // Extract the plain, unformatted text using the Range.Text property.
        string plainText = doc.Range.Text.Trim();

        // Display the extracted text on the console.
        Console.WriteLine(plainText);

        // Optionally, save the extracted text to a separate .txt file.
        File.WriteAllText("output.txt", plainText);
    }
}
