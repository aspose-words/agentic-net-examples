using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a temporary working folder.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a simple PDF file (will be converted to DOCX later).
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(workDir, "source.pdf");
        Document pdfSourceDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSourceDoc);
        pdfBuilder.Writeln("This is the original PDF‑converted document.");
        pdfSourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF and save it as DOCX (simulating a PDF‑to‑DOCX conversion).
        // -----------------------------------------------------------------
        string pdfConvertedDocxPath = Path.Combine(workDir, "pdf_converted.docx");
        Document pdfConvertedDoc = new Document(pdfPath); // Loads PDF.
        pdfConvertedDoc.Save(pdfConvertedDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Create a mail‑merge template DOCX with proper MERGEFIELD fields.
        // -----------------------------------------------------------------
        string mailMergeTemplatePath = Path.Combine(workDir, "mail_merge_template.docx");
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        // Insert MERGEFIELD fields using the InsertField API.
        templateBuilder.Write("Dear ");
        templateBuilder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        templateBuilder.Write(" ");
        templateBuilder.InsertField("MERGEFIELD LastName", "<LastName>");
        templateBuilder.Writeln(",");
        templateBuilder.Write("Your order ");
        templateBuilder.InsertField("MERGEFIELD OrderId", "<OrderId>");
        templateBuilder.Writeln(" has been shipped.");
        templateDoc.Save(mailMergeTemplatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Execute mail merge to produce a merged DOCX.
        // -----------------------------------------------------------------
        string mergedDocxPath = Path.Combine(workDir, "merged_mail_merge.docx");
        Document mailMergeDoc = new Document(mailMergeTemplatePath);
        // Prepare data for the merge.
        DataTable table = new DataTable("Customers");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("OrderId");
        table.Rows.Add("John", "Doe", "12345");
        // Perform the merge.
        mailMergeDoc.MailMerge.Execute(table);
        mailMergeDoc.Save(mergedDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Append the mail‑merged document to the PDF‑converted DOCX,
        //    preserving the destination (PDF‑converted) styles.
        // -----------------------------------------------------------------
        Document destinationDoc = new Document(pdfConvertedDocxPath);
        Document sourceDoc = new Document(mergedDocxPath);
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);
        string finalDocxPath = Path.Combine(workDir, "final_combined.docx");
        destinationDoc.Save(finalDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 6. Validation: ensure the final file exists and contains text from both parts.
        // -----------------------------------------------------------------
        if (!File.Exists(finalDocxPath))
            throw new FileNotFoundException("The combined document was not created.", finalDocxPath);

        string finalText = new Document(finalDocxPath).GetText();
        if (!finalText.Contains("This is the original PDF‑converted document.") ||
            !finalText.Contains("Dear John Doe,"))
            throw new InvalidOperationException("The combined document does not contain expected content.");
    }
}
