using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

// Alias to avoid conflict with System.Range (C# 8+).
using AsposeRange = Aspose.Words.Range;

class ExportRangeToTxt
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document sourceDoc = new Document("InputDocument.docx");

        // Define the range you want to export.
        // Here we export the entire document range; you can adjust as needed.
        AsposeRange exportRange = sourceDoc.Range;

        // Create a new temporary document that contains only the selected range.
        Document rangeDoc = exportRange.ToDocument();

        // Configure TXT save options to preserve line breaks (default ParagraphBreak is CrLf).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Ensure UTF-8 encoding with BOM for proper text representation.
            Encoding = Encoding.UTF8,
            // Preserve page breaks as form feed characters if required.
            // ForcePageBreaks = true,
        };

        // Save the range document as plain text.
        rangeDoc.Save("ExportedRange.txt", txtOptions);
    }
}
