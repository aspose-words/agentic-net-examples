using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Output file names
        const string docPath = "Watermarked.docx";
        const string pdfPath = "Watermarked.pdf";

        // Create a new blank Word document
        Document doc = new Document();

        // Add a text watermark to the document
        doc.Watermark.SetText("Confidential");

        // Save the document as a .docx file (optional, ensures the source file exists)
        doc.Save(docPath);

        // Save the same document directly to PDF format, preserving the watermark
        doc.Save(pdfPath);

        // Simple verification that the PDF file was created
        Console.WriteLine(File.Exists(pdfPath) ? "PDF saved successfully." : "PDF save failed.");
    }
}
