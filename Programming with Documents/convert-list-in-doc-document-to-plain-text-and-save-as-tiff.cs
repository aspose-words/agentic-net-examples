using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ListToTiffConverter
{
    static void Main()
    {
        // Path to the source DOC file containing the list.
        string sourceDocPath = @"C:\Input\ListDocument.doc";

        // Path to the output TIFF file.
        string outputTiffPath = @"C:\Output\ListPlainText.tiff";

        // Load the source document.
        Document sourceDoc = new Document(sourceDocPath);

        // Extract the entire document content as plain text.
        // This includes list items without their formatting.
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document to hold the plain‑text representation.
        Document plainTextDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(plainTextDoc);

        // Write the extracted plain text into the new document.
        builder.Writeln(plainText);

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: set compression (e.g., LZW) to reduce file size.
            TiffCompression = TiffCompression.Lzw,

            // Optional: set resolution (dpi) for higher quality.
            Resolution = 300
        };

        // Save the plain‑text document as a TIFF image.
        plainTextDoc.Save(outputTiffPath, tiffOptions);
    }
}
