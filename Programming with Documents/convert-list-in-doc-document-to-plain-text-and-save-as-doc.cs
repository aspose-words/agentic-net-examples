using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string sourcePath = "Input.doc";

        // Path where the plain‑text DOC will be saved.
        string destinationPath = "Output.doc";

        // Load the original document.
        Document sourceDoc = new Document(sourcePath);

        // Update list labels so they are correct before extracting text.
        sourceDoc.UpdateListLabels();

        // Extract the document's content as plain text.
        // The PlainTextDocument class reads the file and provides the concatenated text.
        PlainTextDocument plainTextDoc = new PlainTextDocument(sourcePath);
        string plainText = plainTextDoc.Text;

        // Create a new blank document and write the extracted plain text into it.
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);
        builder.Writeln(plainText);

        // Save the new document in the DOC format.
        resultDoc.Save(destinationPath, SaveFormat.Doc);
    }
}
