using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input DOCX and the final PDF output.
        const string inputPath = "input.docx";
        const string outputPath = "report.pdf";

        // 1. Create a sample DOCX containing a date placeholder.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Report generated on <<Date>>.");
        builder.Writeln("Additional line with the same placeholder: <<Date>>.");
        source.Save(inputPath, SaveFormat.Docx);

        // 2. Load the DOCX and replace all placeholders with the current date.
        Document doc = new Document(inputPath);
        string placeholder = "<<Date>>";
        string currentDate = DateTime.Now.ToString("D"); // e.g., "Monday, 30 August 2026"
        doc.Range.Replace(placeholder, currentDate, new FindReplaceOptions());

        // 3. Export the updated document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // 4. Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
