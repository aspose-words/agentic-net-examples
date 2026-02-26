using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextConverter
{
    static void Main()
    {
        // Paths to the source DOC file and the resulting DOC file.
        string sourcePath = "input.doc";
        string resultPath = "output.doc";

        // Load the original document (lifecycle rule: load).
        Document sourceDoc = new Document(sourcePath);

        // Configure text save options to simplify list labels (optional, but ensures plain‑text list representation).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true,
            // Ensure the output is plain text.
            SaveFormat = SaveFormat.Text
        };

        // Save the document to a memory stream as plain text (lifecycle rule: save).
        using (MemoryStream txtStream = new MemoryStream())
        {
            sourceDoc.Save(txtStream, txtOptions);
            txtStream.Position = 0; // Reset for reading.

            // Read the plain‑text content.
            string plainText = new StreamReader(txtStream, Encoding.UTF8).ReadToEnd();

            // Create a new blank document (lifecycle rule: create).
            Document plainDoc = new Document();

            // Write the extracted plain text into the new document.
            DocumentBuilder builder = new DocumentBuilder(plainDoc);
            builder.Writeln(plainText);

            // Save the new document as a binary DOC file (lifecycle rule: save).
            plainDoc.Save(resultPath, SaveFormat.Doc);
        }
    }
}
