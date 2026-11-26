using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths
        string docxPath = "Sample.docx";
        string odtPath = "Sample.odt";
        string outputPdfPath = "Combined.pdf";

        // Create a DOCX document with sample content
        Document docxDocument = new Document();
        DocumentBuilder docxBuilder = new DocumentBuilder(docxDocument);
        docxBuilder.Writeln("This is content from the DOCX document.");
        docxDocument.Save(docxPath, SaveFormat.Docx);

        // Create an ODT document with sample content
        Document odtDocument = new Document();
        DocumentBuilder odtBuilder = new DocumentBuilder(odtDocument);
        odtBuilder.Writeln("This is content from the ODT document.");
        odtDocument.Save(odtPath, SaveFormat.Odt);

        // Load the created documents
        Document destination = new Document(docxPath);
        Document source = new Document(odtPath);

        // Append the ODT document to the DOCX document, preserving source formatting
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF
        destination.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created and is not empty
        if (!File.Exists(outputPdfPath))
        {
            throw new FileNotFoundException("The combined PDF file was not created.", outputPdfPath);
        }

        FileInfo pdfInfo = new FileInfo(outputPdfPath);
        if (pdfInfo.Length == 0)
        {
            throw new InvalidOperationException("The combined PDF file is empty.");
        }

        // Optional: output a simple confirmation (no interactive input)
        Console.WriteLine($"Combined PDF created successfully at '{Path.GetFullPath(outputPdfPath)}'.");
    }
}
