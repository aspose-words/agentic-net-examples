using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextJpeg
{
    static void Main()
    {
        // Paths to the source DOC document and the resulting JPEG image.
        string sourcePath = "input.docx";
        string outputPath = "output.jpg";

        // Load the original document (lifecycle rule: create/load).
        Document sourceDoc = new Document(sourcePath);

        // Extract the document's plain text using the PlainTextDocument helper (rule exists).
        PlainTextDocument plainTextDoc = new PlainTextDocument(sourcePath);
        string plainText = plainTextDoc.Text;

        // Create a new blank document and write the extracted plain text into it.
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);
        builder.Writeln(plainText);

        // Prepare image save options for JPEG format (rule exists).
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        jpegOptions.JpegQuality = 100; // optional: set high quality

        // Save the document as a JPEG image (lifecycle rule: save with options).
        resultDoc.Save(outputPath, jpegOptions);
    }
}
