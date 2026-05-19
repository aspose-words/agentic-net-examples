using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for intermediate and final files.
        string templatePath = Path.Combine(outputDir, "MailMergeTemplate.docx");
        string mergedMailDocPath = Path.Combine(outputDir, "MergedMailDoc.docx");
        string pdfConvertedDocPath = Path.Combine(outputDir, "PdfConverted.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // -------------------------------------------------
        // 1. Create a simple mail‑merge template document.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.InsertField("MERGEFIELD Name", "<Name>");
        tmplBuilder.Writeln();
        tmplBuilder.InsertField("MERGEFIELD Address", "<Address>");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Execute mail merge with sample data.
        // -------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("Name");
        data.Columns.Add("Address");
        data.Rows.Add("John Doe", "123 Main St, Anytown");
        data.Rows.Add("Jane Smith", "456 Oak Ave, Othertown");

        Document mergedDoc = new Document(templatePath);
        mergedDoc.MailMerge.Execute(data);
        mergedDoc.Save(mergedMailDocPath);

        // -------------------------------------------------
        // 3. Simulate a PDF‑converted DOCX (source document).
        // -------------------------------------------------
        Document pdfConvertedDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        // Apply a distinct style to demonstrate style preservation.
        pdfBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        pdfBuilder.Writeln("Content originating from a PDF‑converted document.");
        pdfConvertedDoc.Save(pdfConvertedDocPath);

        // -------------------------------------------------
        // 4. Load destination (PDF‑converted) and source (mail‑merged) documents.
        // -------------------------------------------------
        Document destination = new Document(pdfConvertedDocPath);
        Document source = new Document(mergedMailDocPath);

        // -------------------------------------------------
        // 5. Append the mail‑merged document to the destination,
        //    preserving the destination's styles.
        // -------------------------------------------------
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // -------------------------------------------------
        // 6. Save the combined document.
        // -------------------------------------------------
        destination.Save(resultPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 7. Validation – ensure the file exists and contains
        //    expected content from both parts.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        Document resultDoc = new Document(resultPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("Content originating from a PDF‑converted document."))
            throw new InvalidOperationException("Destination content missing in the result.");

        if (!resultText.Contains("John Doe") || !resultText.Contains("Jane Smith"))
            throw new InvalidOperationException("Mail‑merged content missing in the result.");

        // The program finishes without interactive prompts.
    }
}
