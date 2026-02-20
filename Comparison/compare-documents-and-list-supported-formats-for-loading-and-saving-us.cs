using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the two documents to compare.
        string sourcePath = @"C:\Docs\Source.docx";
        string targetPath = @"C:\Docs\Target.docx";

        // Load the documents (format is detected automatically).
        Document sourceDoc = new Document(sourcePath);
        Document targetDoc = new Document(targetPath);

        // Compare the documents. The result (revisions) will be stored in sourceDoc.
        CompareOptions compareOptions = new CompareOptions();
        sourceDoc.Compare(targetDoc, "Comparer", DateTime.Now, compareOptions);

        // Output the number of revisions found.
        Console.WriteLine($"Revisions count after comparison: {sourceDoc.Revisions.Count}");

        // --------------------------------------------------------------------
        // List all load formats that can be saved as DOCX.
        // --------------------------------------------------------------------
        List<string> loadToDocx = new List<string>();
        foreach (LoadFormat loadFormat in Enum.GetValues(typeof(LoadFormat)))
        {
            // Convert the load format to a save format, if possible.
            SaveFormat saveFormat = FileFormatUtil.LoadFormatToSaveFormat(loadFormat);
            if (saveFormat == SaveFormat.Docx)
            {
                // Get a friendly file extension for the load format.
                string extension = FileFormatUtil.LoadFormatToExtension(loadFormat);
                loadToDocx.Add($"{loadFormat} ({extension})");
            }
        }

        Console.WriteLine("\nLoad formats that can be saved as DOCX:");
        foreach (string entry in loadToDocx)
            Console.WriteLine($"- {entry}");

        // --------------------------------------------------------------------
        // List all save formats that can be loaded from DOCX.
        // --------------------------------------------------------------------
        List<string> saveFromDocx = new List<string>();
        foreach (SaveFormat saveFormat in Enum.GetValues(typeof(SaveFormat)))
        {
            // Convert the save format to a load format, if possible.
            LoadFormat loadFormat = FileFormatUtil.SaveFormatToLoadFormat(saveFormat);
            if (loadFormat == LoadFormat.Docx)
            {
                // Get a friendly file extension for the save format.
                string extension = FileFormatUtil.SaveFormatToExtension(saveFormat);
                saveFromDocx.Add($"{saveFormat} ({extension})");
            }
        }

        Console.WriteLine("\nSave formats that can be loaded from DOCX:");
        foreach (string entry in saveFromDocx)
            Console.WriteLine($"- {entry}");
    }
}
