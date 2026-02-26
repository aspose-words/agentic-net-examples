using System;
using Aspose.Words;

class ConvertListToPlainTextRtf
{
    static void Main()
    {
        // Path to the source DOC document that contains a list.
        string sourcePath = "input.doc";

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure list labels are up‑to‑date before extracting plain text.
        sourceDoc.UpdateListLabels();

        // Extract the document's content as plain text (lists are represented as text).
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document.
        Document rtfDoc = new Document();

        // Insert the extracted plain text into the new document.
        DocumentBuilder builder = new DocumentBuilder(rtfDoc);
        builder.Writeln(plainText);

        // Save the new document as RTF.
        rtfDoc.Save("output.rtf", SaveFormat.Rtf);
    }
}
