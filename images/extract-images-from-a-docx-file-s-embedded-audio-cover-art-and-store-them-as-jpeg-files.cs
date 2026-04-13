using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractAudioCoverArt
{
    public static void Main()
    {
        // Prepare deterministic file names and paths.
        string workDir = Directory.GetCurrentDirectory();
        string docPath = Path.Combine(workDir, "SampleWithAudio.docx");
        string audioPath = Path.Combine(workDir, "sample.wav");
        string coverImagePath = Path.Combine(workDir, "cover.png");

        // -----------------------------------------------------------------
        // 1. Create a sample cover image (PNG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        try
        {
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkBlue))
                {
                    g.FillEllipse(brush, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }
            finally
            {
                g.Dispose();
            }

            // Save as PNG – this will be used as the OLE icon (cover art).
            bitmap.Save(coverImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        // -----------------------------------------------------------------
        // 2. Create a dummy audio file (WAV) – just a few bytes.
        // -----------------------------------------------------------------
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
            0x00,0x08,0x00,0x00  // Subchunk2Size
        };
        File.WriteAllBytes(audioPath, wavHeader);

        // -----------------------------------------------------------------
        // 3. Build a DOCX containing the audio OLE object with the cover image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object; the icon image will be the cover art.
        using (FileStream iconStream = new FileStream(coverImagePath, FileMode.Open, FileAccess.Read))
        {
            // Parameters: fileName, isLinked (false = embed), asIcon (true), presentation (icon stream)
            builder.InsertOleObject(audioPath, false, true, iconStream);
        }

        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the document and extract images from OLE icons (cover art).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Save the image data to a memory stream.
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading.

                    // Load the image with Aspose.Drawing and re‑save as JPEG.
                    using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(imgStream))
                    {
                        string outFile = Path.Combine(workDir, $"CoverArt_{extractedCount}.jpg");
                        bmp.Save(outFile, Aspose.Drawing.Imaging.ImageFormat.Jpeg);

                        if (!File.Exists(outFile))
                            throw new InvalidOperationException($"Failed to create image file: {outFile}");

                        extractedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No cover art images were extracted from the document.");
    }
}
