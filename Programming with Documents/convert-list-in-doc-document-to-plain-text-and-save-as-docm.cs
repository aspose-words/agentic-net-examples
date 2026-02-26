using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextConverter
{
    static void Main()
    {
        // Paths to the source DOC file and the destination DOCM file.
        string sourcePath = "input.doc";
        string destinationPath = "output.docm";

        // Load the existing DOC document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure that list labels are up‑to‑date before extracting plain text.
        sourceDoc.UpdateListLabels();

        // Extract the document's plain‑text representation (lists are rendered as text).
        PlainTextDocument plainText = new PlainTextDocument(sourcePath);
        string textContent = plainText.Text;

        // Create a new blank document and write the extracted plain text into it.
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);
        builder.Writeln(textContent);

        // Save the result as a macro‑enabled DOCM file.
        resultDoc.Save(destinationPath, SaveFormat.Docm);
    }
}
