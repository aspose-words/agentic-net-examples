using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file names.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        string docxPath = Path.Combine(outputFolder, "Sample.docx");
        string pdfPath = Path.Combine(outputFolder, "Sample.pdf");

        // -----------------------------------------------------------------
        // Step 1: Create a blank DOCX file (bootstrap input).
        // -----------------------------------------------------------------
        Document blankDoc = new Document();
        blankDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Failed to create the DOCX file.", docxPath);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX, add header/footer, and some body text.
        // -----------------------------------------------------------------
        Document doc = new Document(docxPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("My Header");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("My Footer");

        // Return to the main document body and add a paragraph.
        builder.MoveToSection(0);
        builder.Writeln("Hello World!");

        // -----------------------------------------------------------------
        // Step 3: Convert the document to PDF.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF file.", pdfPath);

        // Optional: Inform the user (no interactive pause).
        Console.WriteLine("DOCX and PDF files have been created successfully at:");
        Console.WriteLine(outputFolder);
    }
}
