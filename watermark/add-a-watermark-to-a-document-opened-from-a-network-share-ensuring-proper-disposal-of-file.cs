using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder that simulates a network share.
        string networkShareFolder = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkShareFolder);

        // Paths for the source and output documents on the simulated network share.
        string sourceDocPath = Path.Combine(networkShareFolder, "Source.docx");
        string outputDocPath = Path.Combine(networkShareFolder, "Watermarked.docx");

        // -----------------------------------------------------------------
        // 1. Create a blank document and save it to the network share.
        // -----------------------------------------------------------------
        var blankDoc = new Document();
        // Use a FileStream inside a using block to ensure the file handle is released.
        using (FileStream createStream = File.Create(sourceDocPath))
        {
            blankDoc.Save(createStream, SaveFormat.Docx);
        }

        // -----------------------------------------------------------------
        // 2. Load the document from the network share.
        // -----------------------------------------------------------------
        Document loadedDoc;
        using (FileStream readStream = File.OpenRead(sourceDocPath))
        {
            loadedDoc = new Document(readStream);
        }

        // -----------------------------------------------------------------
        // 3. Add a text watermark to the loaded document.
        // -----------------------------------------------------------------
        loadedDoc.Watermark.SetText("Confidential");

        // -----------------------------------------------------------------
        // 4. Save the watermarked document back to the network share.
        // -----------------------------------------------------------------
        using (FileStream writeStream = File.Create(outputDocPath))
        {
            loadedDoc.Save(writeStream, SaveFormat.Docx);
        }

        // -----------------------------------------------------------------
        // 5. Simple validation that the output file exists.
        // -----------------------------------------------------------------
        if (File.Exists(outputDocPath))
        {
            Console.WriteLine("Watermark applied successfully. Output file: " + outputDocPath);
        }
        else
        {
            Console.WriteLine("Failed to create the watermarked document.");
        }
    }
}
