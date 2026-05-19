using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for the output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Build a simple document that contains characters which form ligatures (fi, fl, ffi).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Times New Roman"; // A font that supports standard ligatures.
        builder.Font.Size = 48;

        builder.Writeln("Office");   // Contains "ff" and "fi".
        builder.Writeln("Affix");    // Contains "ff" and "fi".
        builder.Writeln("Fluff");    // Contains "fl" and "ff".

        // Render the document to PDF. The default rendering pipeline preserves OpenType
        // features such as ligatures when the chosen font supports them.
        string pdfPath = Path.Combine(artifactsDir, "OpenTypePreserved.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Validate that the PDF file was created and is not empty.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        if (new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("The PDF file is empty.");
    }
}
