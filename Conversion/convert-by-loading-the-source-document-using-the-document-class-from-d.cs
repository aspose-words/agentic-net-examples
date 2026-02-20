using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Create LoadOptions and explicitly set the format to DOCX.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docx,
            // Optional: specify recovery mode to attempt to load corrupted files.
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };

        try
        {
            // Load the document using the Document constructor that accepts a file name and LoadOptions.
            Document doc = new Document(sourcePath, loadOptions);

            // The document is now loaded and can be processed further.
            Console.WriteLine("Document loaded successfully. Page count: " + doc.PageCount);
        }
        catch (UnsupportedFileFormatException ex)
        {
            Console.WriteLine("The file format is not supported: " + ex.Message);
        }
        catch (FileCorruptedException ex)
        {
            Console.WriteLine("The file appears to be corrupted: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
