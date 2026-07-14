using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;   // Required for FindReplaceOptions

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with placeholders.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report for {{Name}}");
        builder.Writeln("Date: {{Date}}");

        const string docxPath = "template.docx";
        template.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX.
        Document doc = new Document(docxPath);

        // Replace placeholders with actual values.
        doc.Range.Replace("{{Name}}", "John Doe", new FindReplaceOptions());
        doc.Range.Replace("{{Date}}", DateTime.Today.ToString("yyyy-MM-dd"), new FindReplaceOptions());

        // Export to PDF.
        const string pdfPath = "report.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
