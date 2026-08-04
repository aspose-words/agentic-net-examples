using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a dummy audio file (WAV header + silence).
        string audioPath = Path.Combine(artifactsDir, "sample.wav");
        byte[] wavHeader = new byte[]
        {
            0x52,0x49,0x46,0x46, // "RIFF"
            0x24,0x08,0x00,0x00, // Chunk size
            0x57,0x41,0x56,0x45, // "WAVE"
            0x66,0x6D,0x74,0x20, // "fmt "
            0x10,0x00,0x00,0x00, // Subchunk1Size
            0x01,0x00,0x01,0x00, // AudioFormat, NumChannels
            0x40,0x1F,0x00,0x00, // SampleRate (8000)
            0x80,0x3E,0x00,0x00, // ByteRate
            0x02,0x00,0x10,0x00, // BlockAlign, BitsPerSample
            0x64,0x61,0x74,0x61, // "data"
            0x00,0x08,0x00,0x00  // Subchunk2Size (2048 bytes of silence)
        };
        File.WriteAllBytes(audioPath, wavHeader);

        // 2. Create a Word document and embed the audio as an OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the audio file as an embedded OLE object (not as an icon).
        // Overload: InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation)
        builder.InsertOleObject(audioPath, false, false, null);
        // Save the document.
        string docPath = Path.Combine(artifactsDir, "AudioDoc.docx");
        doc.Save(docPath);

        // 3. Load the document (optional, we can reuse the same instance).
        Document loadedDoc = new Document(docPath);

        // 4. Extract images (icons) from OLE objects that represent audio tracks.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int thumbIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Consider only shapes that are OLE objects and have an associated image.
            if (shape.OleFormat != null && shape.HasImage)
            {
                // Save the original image to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0;

                    // Load the image using Aspose.Drawing.
                    using (Aspose.Drawing.Image originalImage = Aspose.Drawing.Image.FromStream(imageStream))
                    {
                        const int thumbWidth = 100;
                        const int thumbHeight = 100;

                        // Create a new bitmap for the thumbnail.
                        using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                        {
                            using (Graphics graphics = Graphics.FromImage(thumbBitmap))
                            {
                                // Clear background to white.
                                graphics.Clear(Color.White);
                                // Draw the original image scaled to fit the thumbnail rectangle.
                                graphics.DrawImage(originalImage, new Rectangle(0, 0, thumbWidth, thumbHeight));
                            }

                            // Save the thumbnail as JPEG.
                            string thumbPath = Path.Combine(artifactsDir, $"audio_thumb_{thumbIndex}.jpg");
                            thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);
                            thumbIndex++;
                        }
                    }
                }
            }
        }

        // 5. Validate that at least one thumbnail was created.
        if (thumbIndex == 0)
            throw new InvalidOperationException("No audio track images were found in the document.");
    }
}
