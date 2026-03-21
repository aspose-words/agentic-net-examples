using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderKeywordReplace
{
    static void Main()
    {
        // The keyword that must be present in a header for the replacement to occur.
        const string headerKeyword = "CONFIDENTIAL";

        // Text to find and its replacement.
        const string textToFind = "OldCompany";
        const string textToReplace = "NewCompany";

        // Create a new document with a header that contains the keyword and the text to replace.
        Document doc = new Document();
        Section section = new Section(doc);
        doc.Sections.Add(section);

        // Header that contains the keyword.
        HeaderFooter headerWithKeyword = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(headerWithKeyword);
        Paragraph headerPara = new Paragraph(doc);
        headerWithKeyword.AppendChild(headerPara);
        headerPara.AppendChild(new Run(doc, $"{headerKeyword} - {textToFind}"));

        // Header that does NOT contain the keyword (should remain unchanged).
        HeaderFooter headerWithoutKeyword = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
        section.HeadersFooters.Add(headerWithoutKeyword);
        Paragraph otherHeaderPara = new Paragraph(doc);
        headerWithoutKeyword.AppendChild(otherHeaderPara);
        otherHeaderPara.AppendChild(new Run(doc, $"OtherHeader - {textToFind}"));

        // Iterate through all sections and their headers.
        foreach (Section sec in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in sec.HeadersFooters)
            {
                // Process only headers (not footers) that contain the specific keyword.
                if (headerFooter.IsHeader && headerFooter.GetText().Contains(headerKeyword))
                {
                    FindReplaceOptions options = new FindReplaceOptions
                    {
                        MatchCase = false,
                        FindWholeWordsOnly = false
                    };

                    // Perform the replacement within this header's range.
                    headerFooter.Range.Replace(textToFind, textToReplace, options);
                }
            }
        }

        // Save the modified document to a temporary location.
        string outputPath = Path.Combine(Path.GetTempPath(), "HeaderKeywordReplace_Output.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
