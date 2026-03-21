using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Use paths relative to the executable directory
        string baseDir = AppContext.BaseDirectory;
        string inputFile = Path.Combine(baseDir, "Macro.docm");
        string outputFile = Path.Combine(baseDir, "MacroRelative.docm");

        // Ensure the input file exists; if not, create a minimal DOCM file
        if (!File.Exists(inputFile))
        {
            // Create a new blank document and save it as a macro‑enabled file
            var doc = new Document();
            doc.Save(inputFile, SaveFormat.Docm);
        }

        // Load the document (lifecycle rule: load)
        Document document = new Document(inputFile);

        // Proceed only if the document actually contains VBA macros
        if (document.HasMacros && document.VbaProject != null)
        {
            // Iterate through each VBA module in the project
            foreach (VbaModule module in document.VbaProject.Modules)
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
        document.Save(outputFile);
        Console.WriteLine($"Processed document saved to: {outputFile}");
    }
}
