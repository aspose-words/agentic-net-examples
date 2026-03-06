using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class InsertContentIntoRange
{
    static void Main()
    {
        // Load the source DOCX document.
        // The Document constructor loads the file from the specified path.
        Document doc = new Document("InputDocument.docx");

        // Define the placeholder text that will be replaced.
        // This placeholder should exist somewhere in the document, e.g. "[PLACEHOLDER]".
        string placeholder = "[PLACEHOLDER]";

        // Define the new content that will replace the placeholder.
        string newContent = "This is the inserted content.";

        // Use the Range.Replace method to replace all occurrences of the placeholder
        // with the new content. This operates on the whole document range.
        int replacementsMade = doc.Range.Replace(placeholder, newContent);

        // Optionally, you can verify that a replacement was performed.
        Console.WriteLine($"Replacements made: {replacementsMade}");

        // Save the modified document to a new file.
        // The Save method automatically determines the format from the file extension.
        doc.Save("OutputDocument.docx");
    }
}
