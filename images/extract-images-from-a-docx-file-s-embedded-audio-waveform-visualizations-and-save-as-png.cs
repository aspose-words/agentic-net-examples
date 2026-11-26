using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string audioFilePath = "sample.wav";
        const string docPath = "sample.docx";
        const string waveformImagePath = "waveform.png";

        // 1. Create a minimal WAV file (header only) to use as audio source.
        CreateDummyWav(audioFilePath);

        // 2. Create a sample waveform image that will represent the audio visualization.
        CreateSampleWaveformImage(waveformImagePath);

        // 3. Create a Word document and insert the waveform image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // InsertImage inserts the image as a Shape with HasImage = true.
        builder.InsertImage(waveformImagePath);
        doc.Save(docPath);

        // 4. Reload the document for extraction.
        Document loadedDoc = new Document(docPath);

        // 5. Iterate through all Shape nodes and extract images (waveform visualizations).
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"extracted-image-{imageIndex}.png";
                shape.ImageData.Save(imageFileName);
                if (!File.Exists(imageFileName))
                    throw new InvalidOperationException($"Failed to save image '{imageFileName}'.");
                imageIndex++;
            }
        }

        // 6. Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional cleanup.
        // File.Delete(audioFilePath);
        // File.Delete(docPath);
        // File.Delete(waveformImagePath);
    }

    // Helper method to create a minimal valid WAV file (44‑byte header with no data).
    private static void CreateDummyWav(string path)
    {
        byte[] header = new byte[44];
        // ChunkID "RIFF"
        header[0] = (byte)'R'; header[1] = (byte)'I'; header[2] = (byte)'F'; header[3] = (byte)'F';
        // ChunkSize (36 + SubChunk2Size) = 36
        BitConverter.GetBytes(36).CopyTo(header, 4);
        // Format "WAVE"
        header[8] = (byte)'W'; header[9] = (byte)'A'; header[10] = (byte)'V'; header[11] = (byte)'E';
        // Subchunk1ID "fmt "
        header[12] = (byte)'f'; header[13] = (byte)'m'; header[14] = (byte)'t'; header[15] = (byte)' ';
        // Subchunk1Size (16 for PCM)
        BitConverter.GetBytes(16).CopyTo(header, 16);
        // AudioFormat (1 = PCM)
        BitConverter.GetBytes((short)1).CopyTo(header, 20);
        // NumChannels (1)
        BitConverter.GetBytes((short)1).CopyTo(header, 22);
        // SampleRate (44100)
        BitConverter.GetBytes(44100).CopyTo(header, 24);
        // ByteRate = SampleRate * NumChannels * BitsPerSample/8
        BitConverter.GetBytes(44100 * 1 * 16 / 8).CopyTo(header, 28);
        // BlockAlign = NumChannels * BitsPerSample/8
        BitConverter.GetBytes((short)(1 * 16 / 8)).CopyTo(header, 32);
        // BitsPerSample (16)
        BitConverter.GetBytes((short)16).CopyTo(header, 34);
        // Subchunk2ID "data"
        header[36] = (byte)'d'; header[37] = (byte)'a'; header[38] = (byte)'t'; header[39] = (byte)'a';
        // Subchunk2Size (0)
        BitConverter.GetBytes(0).CopyTo(header, 40);

        using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            fs.Write(header, 0, header.Length);
        }
    }

    // Helper method to create a deterministic sample waveform image.
    private static void CreateSampleWaveformImage(string path)
    {
        const int width = 400;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);

                // Draw a simple waveform using black lines.
                Pen pen = new Pen(Color.Black, 2);
                int midY = height / 2;
                for (int x = 0; x < width; x += 10)
                {
                    int amplitude = (x / 10) % 2 == 0 ? 30 : -30;
                    g.DrawLine(pen, x, midY, x, midY + amplitude);
                }
                pen.Dispose();
            }

            // Save as PNG.
            bitmap.Save(path);
        }

        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample waveform image '{path}'.");
    }
}
