using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a mail‑merge template document (DOCX) with merge fields.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Dear ");
        templateBuilder.InsertField("MERGEFIELD Name", "<Name>");
        templateBuilder.Writeln(",");
        templateBuilder.Writeln("Your address is:");
        templateBuilder.InsertField("MERGEFIELD Address", "<Address>");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 2. Execute mail merge to produce a merged DOCX document.
        // ---------------------------------------------------------------
        // Prepare simple data source.
        DataTable data = new DataTable();
        data.Columns.Add("Name");
        data.Columns.Add("Address");
        data.Rows.Add("John Doe", "123 Main St, Anytown");

        // Load the template and perform the merge.
        Document mergedDoc = new Document(templatePath);
        mergedDoc.MailMerge.Execute(data);
        string mergedDocPath = Path.Combine(outputDir, "MergedFromMailMerge.docx");
        mergedDoc.Save(mergedDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 3. Create a source PDF document and then load it as a DOCX.
        // ---------------------------------------------------------------
        // Create a simple document that will be saved as PDF.
        Document pdfSource = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSource);
        pdfBuilder.Writeln("This is the original PDF‑converted content.");
        string pdfPath = Path.Combine(outputDir, "Source.pdf");
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF; Aspose.Words converts it to an internal Word document.
        Document destinationDoc = new Document(pdfPath);

        // ---------------------------------------------------------------
        // 4. Append the mail‑merged document to the PDF‑converted document.
        //    Preserve destination styles by using UseDestinationStyles.
        // ---------------------------------------------------------------
        destinationDoc.AppendDocument(mergedDoc, ImportFormatMode.UseDestinationStyles);

        // ---------------------------------------------------------------
        // 5. Save the final combined document.
        // ---------------------------------------------------------------
        string finalDocPath = Path.Combine(outputDir, "FinalCombined.docx");
        destinationDoc.Save(finalDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 6. Validation: ensure the file exists and contains expected text.
        // ---------------------------------------------------------------
        if (!File.Exists(finalDocPath))
            throw new InvalidOperationException("The final document was not created.");

        Document validationDoc = new Document(finalDocPath);
        string text = validationDoc.GetText();

        if (!text.Contains("John Doe") || !text.Contains("This is the original PDF‑converted content."))
            throw new InvalidOperationException("The final document does not contain expected content.");

        // The example finishes without requiring any user interaction.
    }
}
