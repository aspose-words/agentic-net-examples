using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the source macro‑enabled document (hard‑coded absolute path)
        string inputFile = @"C:\Docs\Macro.docm";

        // Path where the modified document will be saved
        string outputFile = @"C:\Docs\MacroRelative.docm";

        // Load the document (lifecycle rule: load)
        Document doc = new Document(inputFile);

        // Proceed only if the document actually contains VBA macros
        if (doc.HasMacros && doc.VbaProject != null)
        {
            // Iterate through each VBA module in the project
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Original macro source code
                string source = module.SourceCode;

                // Regular expression that matches absolute Windows file paths
                // Example match: C:\Folder\SubFolder\File.docx
                string pattern = @"[A-Za-z]:\\[^\s""]+";

                // Replace each absolute path with its file name (relative path)
                string updated = Regex.Replace(source, pattern, match =>
                {
                    // Extract only the file name from the full path
                    string fileName = Path.GetFileName(match.Value);
                    return fileName; // Return the relative path
                });

                // Write the modified source back to the module
                module.SourceCode = updated;
            }
        }

        // Save the document with updated macro source code (lifecycle rule: save)
        doc.Save(outputFile);
    }
}
