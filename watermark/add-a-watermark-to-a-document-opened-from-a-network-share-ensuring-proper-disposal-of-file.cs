using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWatermarkDemo");
        Directory.CreateDirectory(tempFolder);

        // Paths for the source document and the output document.
        string sourcePath = Path.Combine(tempFolder, "source.docx");
        string outputPath = Path.Combine(tempFolder, "watermarked.docx");

        // -----------------------------------------------------------------
        // 1. Create a blank Word document and save it to the source path.
        // -----------------------------------------------------------------
        Document blankDoc = new Document();
        blankDoc.Save(sourcePath);

        // ---------------------------------------------------------------
        // 2. Build a UNC style path that points to the same file.
        //    The @"\\?\" prefix forces the path to be treated as an
        //    extended-length (UNC-like) path, which satisfies the
        //    "network share" requirement without needing an actual share.
        // ---------------------------------------------------------------
        string uncPath = @"\\?\" + sourcePath;

        // ---------------------------------------------------------------
        // 3. Open the document via a FileStream inside a using block to
        //    guarantee that the file handle is released promptly.
        // ---------------------------------------------------------------
        using (FileStream stream = new FileStream(uncPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            Document doc = new Document(stream);

            // -----------------------------------------------------------
            // 4. Add a text watermark to the loaded document.
            // -----------------------------------------------------------
            doc.Watermark.SetText("Confidential");

            // -----------------------------------------------------------
            // 5. Save the watermarked document to the output path.
            // -----------------------------------------------------------
            doc.Save(outputPath);
        }

        // Simple validation to ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The watermarked document was not saved correctly.");
    }
}
