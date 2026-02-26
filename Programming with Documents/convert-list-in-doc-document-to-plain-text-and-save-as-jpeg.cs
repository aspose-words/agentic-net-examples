using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextToJpeg
{
    static void Main()
    {
        // Path to the source DOC document that contains a list.
        string sourceDocPath = @"C:\Input\ListDocument.docx";

        // Path where the resulting JPEG image will be saved.
        string outputJpegPath = @"C:\Output\PlainTextImage.jpg";

        // Load the source document.
        Document sourceDoc = new Document(sourceDocPath);

        // Extract the plain text from the document.
        // GetText returns the visible text without field codes or formatting.
        string plainText = sourceDoc.GetText();

        // Create a new blank document to hold the plain text.
        Document plainTextDoc = new Document();

        // Use DocumentBuilder to write the extracted text into the new document.
        DocumentBuilder builder = new DocumentBuilder(plainTextDoc);
        builder.Writeln(plainText);

        // Prepare image save options for JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Save the document as a JPEG image. Only the first page is rendered,
        // which contains the whole plain‑text content.
        plainTextDoc.Save(outputJpegPath, jpegOptions);
    }
}
