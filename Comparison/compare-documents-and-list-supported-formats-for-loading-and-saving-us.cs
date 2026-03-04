using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparisonAndSupportedFormats
{
    static void Main()
    {
        // Paths to the documents to compare.
        const string docPath1 = @"C:\Docs\Original.docx";
        const string docPath2 = @"C:\Docs\Edited.docx";

        // Load the two DOCX documents.
        Document original = new Document(docPath1);
        Document edited = new Document(docPath2);

        // Compare the documents. Revisions will be added to 'original'.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Save the comparison result as a DOCX file.
        const string comparisonResultPath = @"C:\Docs\ComparisonResult.docx";
        original.Save(comparisonResultPath, SaveFormat.Docx);

        // --------------------------------------------------------------------
        // List load formats that can be converted to DOCX (i.e., can be saved as DOCX).
        Console.WriteLine("Load formats that can be saved as DOCX:");
        foreach (LoadFormat loadFmt in Enum.GetValues(typeof(LoadFormat)))
        {
            // Skip the 'Unknown' and 'Auto' placeholders.
            if (loadFmt == LoadFormat.Unknown || loadFmt == LoadFormat.Auto)
                continue;

            // Convert the load format to a save format, if possible.
            SaveFormat? possibleSave = null;
            try
            {
                possibleSave = FileFormatUtil.LoadFormatToSaveFormat(loadFmt);
            }
            catch
            {
                // Conversion not supported; ignore.
            }

            if (possibleSave.HasValue && possibleSave.Value == SaveFormat.Docx)
            {
                string ext = FileFormatUtil.LoadFormatToExtension(loadFmt);
                Console.WriteLine($"- {loadFmt} (extension: {ext})");
            }
        }

        // --------------------------------------------------------------------
        // List save formats that can be loaded from DOCX (i.e., DOCX can be converted to them).
        Console.WriteLine("\nSave formats that can be loaded from DOCX:");
        foreach (SaveFormat saveFmt in Enum.GetValues(typeof(SaveFormat)))
        {
            // Skip the 'Unknown' placeholder.
            if (saveFmt == SaveFormat.Unknown)
                continue;

            // Convert the save format back to a load format, if possible.
            LoadFormat? possibleLoad = null;
            try
            {
                possibleLoad = FileFormatUtil.SaveFormatToLoadFormat(saveFmt);
            }
            catch
            {
                // Conversion not supported; ignore.
            }

            if (possibleLoad.HasValue && possibleLoad.Value == LoadFormat.Docx)
            {
                string ext = FileFormatUtil.SaveFormatToExtension(saveFmt);
                Console.WriteLine($"- {saveFmt} (extension: {ext})");
            }
        }
    }
}
