using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a mail‑merge template DOCX
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Dear ");
        tmplBuilder.InsertField(" MERGEFIELD Name ", "<Name>");
        tmplBuilder.Writeln(",");
        tmplBuilder.Writeln("Your address is:");
        tmplBuilder.InsertField(" MERGEFIELD Address ", "<Address>");
        tmplBuilder.Writeln(".");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Execute mail merge to produce a DOCX document
        // -----------------------------------------------------------------
        // Prepare data source
        DataTable table = new DataTable("Data");
        table.Columns.Add("Name");
        table.Columns.Add("Address");
        table.Rows.Add("John Doe", "123 Main St, Anytown");

        // Load the template and perform the merge
        Document mergedMailDoc = new Document(templatePath);
        mergedMailDoc.MailMerge.Execute(table);
        string mergedMailPath = Path.Combine(outputDir, "MergedFromMailMerge.docx");
        mergedMailDoc.Save(mergedMailPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Create a source DOCX, convert it to PDF, then back to DOCX
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the original document that will be converted from PDF.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // Convert to PDF
        string pdfPath = Path.Combine(outputDir, "Source.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF and save it as a DOCX (PDF‑converted DOCX)
        Document pdfConvertedDoc = new Document(pdfPath);
        string pdfConvertedDocPath = Path.Combine(outputDir, "PdfConverted.docx");
        pdfConvertedDoc.Save(pdfConvertedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Append the mail‑merged DOCX to the PDF‑converted DOCX
        //    preserving destination styles (UseDestinationStyles)
        // -----------------------------------------------------------------
        Document destinationDoc = new Document(pdfConvertedDocPath);
        Document sourceToAppend = new Document(mergedMailPath);
        destinationDoc.AppendDocument(sourceToAppend, ImportFormatMode.UseDestinationStyles);
        string finalPath = Path.Combine(outputDir, "FinalMerged.docx");
        destinationDoc.Save(finalPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Validation
        // -----------------------------------------------------------------
        if (!File.Exists(finalPath))
            throw new InvalidOperationException("The final merged document was not created.");

        // Verify that content from both source documents is present
        string finalText = destinationDoc.GetText();

        if (!finalText.Contains("This is the original document that will be converted from PDF.") ||
            !finalText.Contains("Dear") ||
            !finalText.Contains("John Doe"))
        {
            throw new InvalidOperationException("The final document does not contain expected content from both sources.");
        }

        // Indicate success (no interactive output required)
        Console.WriteLine("Document merging completed successfully.");
    }
}
