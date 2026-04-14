using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a minimal silent WAV file to embed.
        const string audioFileName = "sample.wav";
        CreateSampleWav(audioFileName);

        // Create a new Word document and embed the audio file as an OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the audio file as an embedded OLE object and display it as an icon.
        // Overload: InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation)
        // - isLinked = false (embed the file)
        // - asIcon   = true  (show an icon, which provides a thumbnail image)
        // - presentation = null (use the default icon provided by Aspose.Words)
        builder.InsertOleObject(audioFileName, false, true, null);

        const string docFileName = "AudioDoc.docx";
        doc.Save(docFileName);

        // Extract thumbnail images from shapes that represent the embedded audio.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine a suitable file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = $"AudioThumbnail_{extractedCount}{extension}";

                // Save the image directly using ImageData.Save.
                shape.ImageData.Save(outFile);
                extractedCount++;
            }
        }

        // Validate that at least one thumbnail was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No thumbnail images were extracted from the embedded audio tracks.");
    }

    // Generates a tiny silent WAV file (44‑byte header, no audio data).
    private static void CreateSampleWav(string path)
    {
        byte[] wavHeader = new byte[]
        {
            0x52, 0x49, 0x46, 0x46, 0x24, 0x08, 0x00, 0x00, // "RIFF" + file size
            0x57, 0x41, 0x56, 0x45,                         // "WAVE"
            0x66, 0x6D, 0x74, 0x20,                         // "fmt "
            0x10, 0x00, 0x00, 0x00,                         // Subchunk1Size (16)
            0x01, 0x00,                                     // AudioFormat (PCM)
            0x01, 0x00,                                     // NumChannels
            0x40, 0x1F, 0x00, 0x00,                         // SampleRate (8000)
            0x80, 0x3E, 0x00, 0x00,                         // ByteRate
            0x02, 0x00,                                     // BlockAlign
            0x10, 0x00,                                     // BitsPerSample
            0x64, 0x61, 0x74, 0x61,                         // "data"
            0x00, 0x08, 0x00, 0x00                          // Subchunk2Size (0)
        };
        File.WriteAllBytes(path, wavHeader);
    }
}
