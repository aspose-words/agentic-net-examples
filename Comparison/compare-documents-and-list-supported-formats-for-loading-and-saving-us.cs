using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparisonAndFormatListing
{
    static void Main()
    {
        // Paths to the documents to compare.
        string docPath1 = @"C:\Docs\Original.docx";
        string docPath2 = @"C:\Docs\Edited.docx";

        // Load the two documents (lifecycle: load).
        Document originalDoc = new Document(docPath1);
        Document editedDoc = new Document(docPath2);

        // Compare the documents (produces revisions in the original document).
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

        // Save the comparison result (lifecycle: save).
        string comparisonResultPath = @"C:\Docs\ComparisonResult.docx";
        originalDoc.Save(comparisonResultPath, SaveFormat.Docx);

        // --------------------------------------------------------------------
        // List all load formats supported by Aspose.Words.
        Console.WriteLine("Supported Load Formats:");
        foreach (LoadFormat loadFormat in Enum.GetValues(typeof(LoadFormat)))
        {
            // Skip the Unknown value – it represents an unsupported format.
            if (loadFormat == LoadFormat.Unknown) continue;

            // Display the enum name and its integer value.
            Console.WriteLine($"  {loadFormat} = {(int)loadFormat}");
        }

        // List all save formats supported by Aspose.Words.
        Console.WriteLine("\nSupported Save Formats:");
        foreach (SaveFormat saveFormat in Enum.GetValues(typeof(SaveFormat)))
        {
            // Skip the Unknown value – it represents an invalid format.
            if (saveFormat == SaveFormat.Unknown) continue;

            // Display the enum name and its integer value.
            Console.WriteLine($"  {saveFormat} = {(int)saveFormat}");
        }

        // --------------------------------------------------------------------
        // Demonstrate detection of a file's format using FileFormatUtil.
        // This shows how to discover the load format of a DOCX file.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(comparisonResultPath);
        Console.WriteLine($"\nDetected format of '{Path.GetFileName(comparisonResultPath)}': {formatInfo.LoadFormat}");

        // Example: Convert the detected load format to a save format and save as that format.
        SaveFormat detectedSaveFormat = FileFormatUtil.LoadFormatToSaveFormat(formatInfo.LoadFormat);
        string convertedPath = Path.ChangeExtension(comparisonResultPath, FileFormatUtil.SaveFormatToExtension(detectedSaveFormat));
        originalDoc.Save(convertedPath, detectedSaveFormat);
        Console.WriteLine($"Document also saved as: {convertedPath}");
    }
}
