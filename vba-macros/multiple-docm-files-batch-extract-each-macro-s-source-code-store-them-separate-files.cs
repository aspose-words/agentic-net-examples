using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class MacroExtractor
{
    static void Main()
    {
        // Folder that contains the DOCM files.
        string inputFolder = @"C:\Docs";

        // Folder where extracted macro source files will be saved.
        string outputFolder = @"C:\MacrosOutput";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Process each DOCM file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docm"))
        {
            // Load the Word document.
            Document doc = new Document(docPath);

            // Skip files that do not contain VBA macros.
            if (!doc.HasMacros)
                continue;

            // Access the VBA project and its modules.
            VbaProject vbaProject = doc.VbaProject;
            VbaModuleCollection modules = vbaProject.Modules;

            int moduleIndex = 0;
            foreach (VbaModule module in modules)
            {
                // Retrieve the source code of the current module.
                string sourceCode = module.SourceCode ?? string.Empty;

                // Create a new document that will hold the macro source as plain text.
                Document macroDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(macroDoc);

                // Optional header comment with the module name.
                builder.Writeln($"' Module: {module.Name}");
                builder.Writeln(sourceCode);

                // Build a unique file name for the extracted macro.
                string safeModuleName = string.IsNullOrEmpty(module.Name)
                    ? $"Module{moduleIndex}"
                    : module.Name;

                string outputFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{safeModuleName}.bas";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the macro source as a plain‑text file.
                macroDoc.Save(outputPath);

                moduleIndex++;
            }
        }
    }
}
