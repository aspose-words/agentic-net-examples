using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToTiffConverter
{
    static void Main()
    {
        // Path to the source DOC document containing the list.
        string sourceDocPath = "Input/DocumentWithList.doc";

        // Path where the resulting TIFF image will be saved.
        string outputTiffPath = "Output/ListPlainText.tiff";

        // Load the source document.
        Document sourceDoc = new Document(sourceDocPath);

        // Extract the entire document content as plain text.
        // This includes the list items in textual form.
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document to hold the plain‑text representation.
        Document plainTextDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(plainTextDoc);

        // Write the extracted plain text into the new document.
        builder.Writeln(plainText);

        // Configure image save options to render the document as a TIFF file.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Optional: set compression, resolution, etc.
        // tiffOptions.TiffCompression = TiffCompression.Lzw;
        // tiffOptions.Resolution = 300;

        // Save the plain‑text document as a TIFF image.
        plainTextDoc.Save(outputTiffPath, tiffOptions);
    }
}
