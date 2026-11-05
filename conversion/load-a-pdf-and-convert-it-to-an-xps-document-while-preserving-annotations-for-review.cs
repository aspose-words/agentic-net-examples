using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and final XPS files
        string pdfPath = "sample.pdf";
        string xpsPath = "output.xps";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF (sample input)
        // -----------------------------------------------------------------
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);
        builder.Writeln("Sample PDF content with annotation placeholder.");

        // Save as PDF – this will be the source file for the conversion step
        wordDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // 2. Load the generated PDF and convert it to XPS (preserving content)
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // ---------------------------------------------------------------
        // 3. Validate that the XPS file was created and is not empty
        // ---------------------------------------------------------------
        if (!File.Exists(xpsPath) || new FileInfo(xpsPath).Length == 0)
        {
            throw new InvalidOperationException("XPS conversion failed or the output file is empty.");
        }

        // Optional: clean up temporary PDF (comment out if you need to inspect it)
        // File.Delete(pdfPath);
    }
}
