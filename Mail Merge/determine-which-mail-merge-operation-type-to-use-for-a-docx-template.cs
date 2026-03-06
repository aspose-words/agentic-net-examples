using System;
using Aspose.Words;
using Aspose.Words.Settings;

class MailMergeOperationDetector
{
    static void Main()
    {
        // Path to the DOCX template you want to inspect.
        string docPath = @"C:\Temp\Template.docx";

        // Load the document using Aspose.Words.
        Document doc = new Document(docPath);

        // Retrieve the main document type from the mail‑merge settings.
        MailMergeMainDocumentType mainType = doc.MailMergeSettings.MainDocumentType;

        // Determine the appropriate mail‑merge operation based on the main document type.
        string operation = mainType switch
        {
            MailMergeMainDocumentType.FormLetters => "Form Letters (standard mail‑merge)",
            MailMergeMainDocumentType.MailingLabels => "Mailing Labels",
            MailMergeMainDocumentType.Envelopes => "Envelopes",
            MailMergeMainDocumentType.Catalog => "Catalog",
            MailMergeMainDocumentType.Email => "Email messages",
            MailMergeMainDocumentType.Fax => "Fax documents",
            MailMergeMainDocumentType.NotAMergeDocument => "Not a mail‑merge document",
            _ => "Unknown or default type"
        };

        // Output the result.
        Console.WriteLine($"MainDocumentType: {mainType}");
        Console.WriteLine($"Recommended mail‑merge operation: {operation}");
    }
}
