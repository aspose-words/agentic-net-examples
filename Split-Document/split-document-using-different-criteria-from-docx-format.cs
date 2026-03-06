using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = Path.Combine("Data", "SourceDocument.docx");

            // Load the DOCX document using the Document constructor (lifecycle rule).
            Document doc = new Document(inputPath);

            // -----------------------------------------------------------------
            // Example 1: Split the document into separate HTML files at each
            //            section break.
            // -----------------------------------------------------------------
            HtmlSaveOptions sectionSplitOptions = new HtmlSaveOptions
            {
                // Split criteria – SectionBreak will cause a new HTML part for each
                // section in the source document.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Save the document. The Save method determines the format from the
            // file extension and applies the provided HtmlSaveOptions.
            string sectionOutputPath = Path.Combine("Output", "SectionSplit.html");
            doc.Save(sectionOutputPath, sectionSplitOptions);

            // -----------------------------------------------------------------
            // Example 2: Split the document at heading paragraphs (Heading 1 and
            //            Heading 2) and limit the split to heading level 2.
            // -----------------------------------------------------------------
            HtmlSaveOptions headingSplitOptions = new HtmlSaveOptions
            {
                // Split at paragraphs that use heading styles.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                // Maximum heading level to consider for splitting (1‑9). 2 means
                // Heading 1 and Heading 2 will trigger a split.
                DocumentSplitHeadingLevel = 2
            };

            string headingOutputPath = Path.Combine("Output", "HeadingSplit.html");
            doc.Save(headingOutputPath, headingSplitOptions);

            // -----------------------------------------------------------------
            // Example 3: Split the document at explicit page breaks.
            // -----------------------------------------------------------------
            HtmlSaveOptions pageBreakSplitOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.PageBreak
            };

            string pageBreakOutputPath = Path.Combine("Output", "PageBreakSplit.html");
            doc.Save(pageBreakOutputPath, pageBreakSplitOptions);

            // Inform the user that the operation completed.
            Console.WriteLine("Document split and saved successfully.");
        }
    }
}
