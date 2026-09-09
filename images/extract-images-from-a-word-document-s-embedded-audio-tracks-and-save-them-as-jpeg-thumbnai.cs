using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        const string audioFileName = "sample.wav";
        const string docFileName = "AudioDoc.docx";
        const string outputFolder = "OutputImages";

        Directory.CreateDirectory(outputFolder);

        // 1. Create a dummy WAV file.
        CreateDummyWavFile(audioFileName);

        // 2. Create a document and embed the audio as an OLE object with a custom icon.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a simple 32x32 PNG icon.
        using (MemoryStream iconStream = CreateSampleIcon())
        {
            // Insert the audio file as an OLE object displayed as an icon.
            // Parameters: (fileName, isLinked, asIcon, presentation)
            builder.InsertOleObject(audioFileName, false, true, iconStream);
        }

        doc.Save(docFileName);

        // 3. Load the document and extract the OLE icon images.
        Document loadedDoc = new Document(docFileName);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // The shape must contain an image (the icon) and an OLE package.
            if (!shape.HasImage || shape.OleFormat?.OlePackage == null)
                continue;

            // Determine the file extension of the embedded OLE package.
            string embeddedFileName = shape.OleFormat.OlePackage.FileName;
            string embeddedExtension = Path.GetExtension(embeddedFileName)?.ToLowerInvariant();

            // Consider common audio extensions.
            if (string.IsNullOrEmpty(embeddedExtension) ||
                !(embeddedExtension == ".wav" || embeddedExtension == ".mp3" || embeddedExtension == ".wma"))
                continue;

            // Save the icon image as a JPEG thumbnail.
            string outputImagePath = Path.Combine(outputFolder,
                $"AudioThumbnail_{imageIndex}.jpg");

            shape.ImageData.Save(outputImagePath);

            if (!File.Exists(outputImagePath))
                throw new InvalidOperationException($"Failed to create thumbnail: {outputImagePath}");

            imageIndex++;
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No audio‑related images were found in the document.");

        // Optional cleanup.
        // File.Delete(audioFileName);
        // File.Delete(docFileName);
    }

    // Creates a minimal WAV file containing silence.
    private static void CreateDummyWavFile(string filePath)
    {
        byte[] wavHeader = new byte[]
        {
            0x52,0x49,0x46,0x46, // "RIFF"
            0x24,0x08,0x00,0x00, // Chunk size (2084 bytes)
            0x57,0x41,0x56,0x45, // "WAVE"
            0x66,0x6D,0x74,0x20, // "fmt "
            0x10,0x00,0x00,0x00, // Subchunk1Size (16)
            0x01,0x00,0x01,0x00, // AudioFormat=1 (PCM), NumChannels=1
            0x40,0x1F,0x00,0x00, // SampleRate=8000
            0x40,0x1F,0x00,0x00, // ByteRate=8000
            0x01,0x00,0x08,0x00, // BlockAlign=1, BitsPerSample=8
            0x64,0x61,0x74,0x61, // "data"
            0x00,0x08,0x00,0x00  // Subchunk2Size (2048 bytes of silence)
        };

        byte[] silence = new byte[2048];

        using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        {
            fs.Write(wavHeader, 0, wavHeader.Length);
            fs.Write(silence, 0, silence.Length);
        }
    }

    // Generates a simple 32x32 PNG icon and returns it as a MemoryStream.
    private static MemoryStream CreateSampleIcon()
    {
        const int size = 32;
        Bitmap bitmap = new Bitmap(size, size);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightGray);
        // Draw a simple "audio" symbol (a rectangle with a small triangle).
        graphics.FillRectangle(Brushes.DarkGray, 8, 8, 16, 16);
        graphics.FillPolygon(Brushes.Black, new[]
        {
            new Point(24, 8), new Point(28, 12), new Point(24, 16)
        });
        graphics.Dispose();

        MemoryStream ms = new MemoryStream();
        bitmap.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
        ms.Position = 0;

        bitmap.Dispose();
        return ms;
    }
}
