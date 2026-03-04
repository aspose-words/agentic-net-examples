using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DotxConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the source DOTX file.
            string inputFolder = @"C:\Docs\Input\";
            // Path to the folder where the converted files will be saved.
            string outputFolder = @"C:\Docs\Output\";

            // Load the DOTX template. The Document constructor automatically detects the format.
            Document template = new Document(inputFolder + "Template.dotx");

            // -----------------------------------------------------------------
            // Example 1: Convert the DOTX template to a DOCX document (macro‑free).
            // -----------------------------------------------------------------
            string docxPath = outputFolder + "Converted.docx";
            template.Save(docxPath, SaveFormat.Docx);

            // -----------------------------------------------------------------
            // Example 2: Convert the DOTX template to a PDF file.
            // -----------------------------------------------------------------
            string pdfPath = outputFolder + "Converted.pdf";
            template.Save(pdfPath, SaveFormat.Pdf);

            // -----------------------------------------------------------------
            // Example 3: Convert the DOTX template to a flat OPC XML file.
            // -----------------------------------------------------------------
            string flatOpcPath = outputFolder + "Converted.flatopc";
            // Use OoxmlSaveOptions to specify the flat OPC format.
            OoxmlSaveOptions flatOpcOptions = new OoxmlSaveOptions(SaveFormat.FlatOpcTemplate);
            template.Save(flatOpcPath, flatOpcOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
