using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains a bookmark named "InsertHere"
            string templatePath = @"C:\Docs\Template.docx";

            // Paths of the documents that will be inserted dynamically
            List<string> documentsToInsert = new List<string>
            {
                @"C:\Docs\Part1.docx",
                @"C:\Docs\Part2.docx",
                @"C:\Docs\Part3.docx"
            };

            // Load the template document (create rule)
            Document template = new Document(templatePath);

            // Create a DocumentBuilder attached to the template (create rule)
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert each document at the bookmark location
            foreach (string docPath in documentsToInsert)
            {
                // Load the source document (create rule)
                Document source = new Document(docPath);

                // Move the cursor to the bookmark where the insertion should occur
                builder.MoveToBookmark("InsertHere");

                // Insert the source document preserving its original formatting
                // (InsertDocument rule)
                builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
            }

            // Optionally, append another document at the end of the result
            Document appendix = new Document(@"C:\Docs\Appendix.docx");
            template.AppendDocument(appendix, ImportFormatMode.UseDestinationStyles); // AppendDocument rule

            // Save the final document (save rule)
            string outputPath = @"C:\Docs\Result.docx";
            template.Save(outputPath, SaveFormat.Docx);
        }
    }
}
