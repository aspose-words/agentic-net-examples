using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a date placeholder.
        Document sample = new Document();
        DocumentBuilder builder = new DocumentBuilder(sample);
        builder.Writeln("Monthly Report");
        builder.Writeln("Generated on <<Date>>.");
        sample.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the DOCX that was just created.
        Document doc = new Document("input.docx");

        // Step 3: Replace all occurrences of the placeholder with the current date.
        string placeholder = "<<Date>>";
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
        doc.Range.Replace(placeholder, currentDate, new FindReplaceOptions());

        // Step 4: Export the updated document to PDF.
        string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Step 5: Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");
    }
}
