using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Prepare temporary files
        string tempDir = Path.Combine(Path.GetTempPath(), "OleExample");
        Directory.CreateDirectory(tempDir);

        // Minimal PDF content
        string pdfPath = Path.Combine(tempDir, "sample.pdf");
        string pdfContent = "%PDF-1.4\n" +
                            "1 0 obj\n" +
                            "<< /Type /Catalog /Pages 2 0 R >>\n" +
                            "endobj\n" +
                            "2 0 obj\n" +
                            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>\n" +
                            "endobj\n" +
                            "3 0 obj\n" +
                            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>\n" +
                            "endobj\n" +
                            "4 0 obj\n" +
                            "<< /Length 44 >>\n" +
                            "stream\n" +
                            "BT\n" +
                            "70 150 Td\n" +
                            "/Helvetica 12 Tf\n" +
                            "(Hello PDF) Tj\n" +
                            "ET\n" +
                            "endstream\n" +
                            "endobj\n" +
                            "5 0 obj\n" +
                            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\n" +
                            "endobj\n" +
                            "xref\n" +
                            "0 6\n" +
                            "0000000000 65535 f \n" +
                            "0000000010 00000 n \n" +
                            "0000000065 00000 n \n" +
                            "0000000116 00000 n \n" +
                            "0000000211 00000 n \n" +
                            "0000000325 00000 n \n" +
                            "trailer\n" +
                            "<< /Root 1 0 R /Size 6 >>\n" +
                            "startxref\n" +
                            "398\n" +
                            "%%EOF";
        File.WriteAllText(pdfPath, pdfContent);

        // Minimal ICO (1x1 black pixel) encoded in base64
        string icoPath = Path.Combine(tempDir, "icon.ico");
        byte[] icoBytes = Convert.FromBase64String(
            "AAABAAEAEBAAAAEAIABoBQAAFgAAACgAAAAQAAAAIAAAAAEAGAAAAAAAAAAA");
        File.WriteAllBytes(icoPath, icoBytes);

        // Start Word via COM late binding
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("Microsoft Word is not installed on this machine.");
            return;
        }

        dynamic wordApp = null;
        dynamic doc = null;
        dynamic range = null;
        dynamic oleShape = null;

        try
        {
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;

            doc = wordApp.Documents.Add();

            range = doc.Range(0, 0);
            oleShape = range.InlineShapes.AddOLEObject(
                ClassType: "AcroExch.Document.DC",
                FileName: pdfPath,
                LinkToFile: false,
                DisplayAsIcon: true,
                IconFileName: icoPath,
                IconLabel: "My PDF Document",
                IconIndex: Type.Missing,
                Range: Type.Missing);

            // Set custom display size (points)
            oleShape.Width = 100f;
            oleShape.Height = 100f;

            // Save the document
            string docPath = Path.Combine(tempDir, "Result.docx");
            doc.SaveAs2(docPath);
        }
        finally
        {
            // Clean up COM objects
            if (oleShape != null) Marshal.FinalReleaseComObject(oleShape);
            if (range != null) Marshal.FinalReleaseComObject(range);
            if (doc != null)
            {
                doc.Close();
                Marshal.FinalReleaseComObject(doc);
            }
            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.FinalReleaseComObject(wordApp);
            }
        }
    }
}
