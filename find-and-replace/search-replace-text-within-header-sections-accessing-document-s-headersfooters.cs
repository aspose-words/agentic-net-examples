using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderReplaceExample
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Ensure the document has at least one section
        Section section = doc.Sections[0];

        // Create a primary header and add placeholder text
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        Paragraph headerParagraph = new Paragraph(doc);
        headerParagraph.AppendChild(new Run(doc, "_FullName_"));
        header.AppendChild(headerParagraph);
        section.HeadersFooters.Add(header);

        // Text to find and its replacement
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Options for the replace operation
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Perform replace in all header types that may exist
        foreach (Section sec in doc.Sections)
        {
            HeaderFooterCollection headersFooters = sec.HeadersFooters;
            ReplaceInHeader(headersFooters, HeaderFooterType.HeaderPrimary, findText, replaceText, options);
            ReplaceInHeader(headersFooters, HeaderFooterType.HeaderFirst,   findText, replaceText, options);
            ReplaceInHeader(headersFooters, HeaderFooterType.HeaderEven,   findText, replaceText, options);
        }

        // Save the modified document to the current directory
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    private static void ReplaceInHeader(HeaderFooterCollection collection,
                                         HeaderFooterType headerType,
                                         string pattern,
                                         string replacement,
                                         FindReplaceOptions options)
    {
        HeaderFooter header = collection[headerType];
        if (header != null)
        {
            header.Range.Replace(pattern, replacement, options);
        }
    }
}
