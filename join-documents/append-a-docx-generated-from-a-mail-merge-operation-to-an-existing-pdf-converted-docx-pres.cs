using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        string pdfPath = Path.Combine(outputDir, "Source.pdf");
        string pdfConvertedPath = Path.Combine(outputDir, "PdfConverted.docx");
        string finalPath = Path.Combine(outputDir, "Final.docx");

        // 1. Create a mail‑merge template with fields
        Document template = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(template);
        templateBuilder.Write("Dear ");
        templateBuilder.InsertField(" MERGEFIELD FirstName ", "<FirstName>");
        templateBuilder.Write(" ");
        templateBuilder.InsertField(" MERGEFIELD LastName ", "<LastName>");
        templateBuilder.Writeln(":");
        templateBuilder.InsertField(" MERGEFIELD Message ", "<Message>");
        template.Save(templatePath, SaveFormat.Docx);

        // 2. Execute mail merge using a DataTable
        DataTable data = new DataTable("Data");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello from mail merge!");
        Document mergedDoc = new Document(templatePath);
        mergedDoc.MailMerge.Execute(data);
        mergedDoc.Save(mergedPath, SaveFormat.Docx);

        // 3. Create a simple PDF document
        Document pdfSource = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSource);
        pdfBuilder.Writeln("This is the content of the original PDF document.");
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // 4. Load the PDF and convert it to DOCX
        Document pdfConverted = new Document(pdfPath);
        pdfConverted.Save(pdfConvertedPath, SaveFormat.Docx);

        // 5. Load destination (PDF‑converted DOCX) and source (mail‑merged DOCX)
        Document dstDoc = new Document(pdfConvertedPath);
        Document srcDoc = new Document(mergedPath);

        // 6. Append source document to destination while preserving destination styles
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // 7. Save the final combined document
        dstDoc.Save(finalPath, SaveFormat.Docx);

        // 8. Validate that the output file exists and contains expected content
        if (!File.Exists(finalPath))
            throw new InvalidOperationException("Final document was not created.");

        string finalText = dstDoc.GetText();
        if (!finalText.Contains("This is the content of the original PDF document.") ||
            !finalText.Contains("Dear John Doe:"))
            throw new InvalidOperationException("Final document does not contain expected content.");
    }
}
