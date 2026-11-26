using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a Word document with a form field.
        Document wordWithForm = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordWithForm);
        builder.Writeln("Please fill the form below:");
        // Insert a text input form field.
        builder.InsertTextInput(
            "TextInput1",
            TextFormFieldType.Regular,
            "",
            "Default value",
            50);

        // Save the document as PDF (this will create PDF form fields).
        string pdfPath = "input.pdf";
        wordWithForm.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF input file was not created.");

        // Step 2: Load the PDF and convert it to DOCX while preserving form fields.
        Document pdfDocument = new Document(pdfPath);
        string docxPath = "output.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX output exists.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX output file was not created.");

        // Indicate success.
        Console.WriteLine("Conversion completed successfully.");
    }
}
