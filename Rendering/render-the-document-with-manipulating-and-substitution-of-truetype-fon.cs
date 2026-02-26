using System;
using Aspose.Words;
using Aspose.Words.Saving;

class FontSubstitutionExample
{
    static void Main()
    {
        // Paths to the source Word document and the destination PDF file.
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Use DocumentBuilder to modify the document's content and fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph using Arial (a TrueType font that can be substituted).
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses Arial.");

        // Write another paragraph using Times New Roman.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses Times New Roman.");

        // Configure PDF save options to replace the specified TrueType fonts
        // with their core PDF Type 1 equivalents.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UseCoreFonts = true // Enables substitution for Arial, Times New Roman, Courier New, Symbol.
        };

        // Save the modified document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}
