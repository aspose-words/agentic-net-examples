using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with a form field.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF with a text input form field:");
        // Insert a legacy text input form field.
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Default value", 50);
        // Save the document as PDF. The form field will be converted to a PDF form field.
        string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF that contains form fields.
        Document pdfDoc = new Document(pdfPath);
        // Convert the PDF to DOCX while preserving form fields.
        string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }

        // Optional: Clean up generated files (comment out if you want to inspect them).
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
