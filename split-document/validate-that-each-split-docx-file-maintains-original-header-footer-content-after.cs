using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace HeaderFooterSplitValidation
{
    class Program
    {
        // Paths used for the demonstration.
        private const string ArtifactsDir = @"./Artifacts/";
        private const string OriginalDocPath = ArtifactsDir + "OriginalDocument.docx";

        static void Main()
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(ArtifactsDir);

            // 1. Create a document with multiple sections, each having identical headers and footers.
            Document originalDoc = CreateDocumentWithHeadersFooters();

            // 2. Save the original document to disk.
            originalDoc.Save(OriginalDocPath);

            // 3. Load the saved document back.
            Document loadedDoc = new Document(OriginalDocPath);

            // 4. Validate that every section retained the original header and footer text.
            ValidateHeadersFooters(loadedDoc);
        }

        /// <summary>
        /// Creates a document that contains three sections.
        /// Each section has a primary header and a primary footer with known text.
        /// </summary>
        private static Document CreateDocumentWithHeadersFooters()
        {
            // Create a blank document.
            Document doc = new Document();

            // Use DocumentBuilder for convenient content creation.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the header text.
            const string headerText = "Sample Header";
            // Define the footer text.
            const string footerText = "Sample Footer";

            // Configure the first section's header and footer.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write(headerText);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write(footerText);

            // Return to the main body of the first section.
            builder.MoveToSection(0);
            builder.Writeln("Content of Section 1.");

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            // The new section inherits the header/footer settings from the previous one.
            builder.Writeln("Content of Section 2.");

            // Insert another section break for the third section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3.");

            return doc;
        }

        /// <summary>
        /// Checks that each section in the provided document contains the expected header and footer text.
        /// Throws an exception if any mismatch is found.
        /// </summary>
        private static void ValidateHeadersFooters(Document doc)
        {
            // Expected texts (must match those used in CreateDocumentWithHeadersFooters).
            const string expectedHeader = "Sample Header";
            const string expectedFooter = "Sample Footer";

            // Iterate over all sections in the document.
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                Section section = doc.Sections[i];
                HeaderFooterCollection hfCollection = section.HeadersFooters;

                // Retrieve primary header and footer; they may be null if not present.
                HeaderFooter header = hfCollection[HeaderFooterType.HeaderPrimary];
                HeaderFooter footer = hfCollection[HeaderFooterType.FooterPrimary];

                // Extract the text, trimming whitespace and line breaks.
                string actualHeader = header?.GetText().Trim() ?? string.Empty;
                string actualFooter = footer?.GetText().Trim() ?? string.Empty;

                // Validate header.
                if (!actualHeader.Equals(expectedHeader, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Header mismatch in section {i + 1}. Expected: \"{expectedHeader}\", Actual: \"{actualHeader}\"");
                }

                // Validate footer.
                if (!actualFooter.Equals(expectedFooter, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Footer mismatch in section {i + 1}. Expected: \"{expectedFooter}\", Actual: \"{actualFooter}\"");
                }
            }

            // If we reach this point, all sections have the correct header/footer.
            Console.WriteLine("All sections retain the original header and footer content.");
        }
    }
}
