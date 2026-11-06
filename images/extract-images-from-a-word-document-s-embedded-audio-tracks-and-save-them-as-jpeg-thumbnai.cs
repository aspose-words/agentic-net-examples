using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a dummy WAV file (header only, no audio data).
        const string audioFilePath = "sample.wav";
        CreateDummyWav(audioFilePath);

        // Create a sample icon image that will be used as the OLE object's display picture.
        const string iconPath = "icon.png";
        CreateSampleIcon(iconPath);

        // Load the WAV and icon files into memory streams.
        using (MemoryStream oleStream = new MemoryStream(File.ReadAllBytes(audioFilePath)))
        using (MemoryStream iconStream = new MemoryStream(File.ReadAllBytes(iconPath)))
        {
            // Reset stream positions to ensure they are read from the beginning.
            oleStream.Position = 0;
            iconStream.Position = 0;

            // Create a new Word document and embed the audio file as an OLE object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the OLE object. The overload expects:
            // (Stream oleObjectData, string progId, bool isLink, Stream iconData)
            builder.InsertOleObject(oleStream, "Package", false, iconStream);

            const string docPath = "sample.docx";
            doc.Save(docPath);
        }

        // Reload the document (simulates typical usage scenario).
        Document loadedDoc = new Document("sample.docx");

        // Extract thumbnail images (icons) from the embedded audio OLE objects.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // The shape must have an image (icon) and be an OLE object.
            if (shape.HasImage && shape.OleFormat != null)
            {
                string imagePath = $"audio-thumbnail-{imageIndex}.jpg";
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }
        }

        // Ensure at least one thumbnail was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No audio thumbnail images were extracted.");

        // Optional cleanup of temporary files.
        // File.Delete(audioFilePath);
        // File.Delete("sample.docx");
        // File.Delete(iconPath);
    }

    // Generates a minimal WAV file with a proper header but no audio data.
    private static void CreateDummyWav(string path)
    {
        byte[] header = new byte[]
        {
            0x52, 0x49, 0x46, 0x46, // "RIFF"
            0x24, 0x00, 0x00, 0x00, // Chunk size (36 + SubChunk2Size)
            0x57, 0x41, 0x56, 0x45, // "WAVE"
            0x66, 0x6D, 0x74, 0x20, // "fmt "
            0x10, 0x00, 0x00, 0x00, // SubChunk1Size (16 for PCM)
            0x01, 0x00,             // AudioFormat (1 = PCM)
            0x01, 0x00,             // NumChannels (1)
            0x40, 0x1F, 0x00, 0x00, // SampleRate (8000 Hz)
            0x80, 0x3E, 0x00, 0x00, // ByteRate (SampleRate * NumChannels * BitsPerSample/8)
            0x02, 0x00,             // BlockAlign (NumChannels * BitsPerSample/8)
            0x10, 0x00,             // BitsPerSample (16)
            0x64, 0x61, 0x74, 0x61, // "data"
            0x00, 0x00, 0x00, 0x00  // SubChunk2Size (0 data bytes)
        };

        using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            fs.Write(header, 0, header.Length);
        }
    }

    // Creates a simple PNG icon image to be used as the OLE object's display picture.
    private static void CreateSampleIcon(string path)
    {
        const int width = 100;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple rectangle to make the icon recognizable.
            g.DrawRectangle(Pens.Black, 10, 10, width - 20, height - 20);
            bitmap.Save(path);
        }
    }
}
