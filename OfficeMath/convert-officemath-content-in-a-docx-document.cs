using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class OfficeMathConversion
{
    static void Main()
    {
        // Path to the input DOCX file that contains OfficeMath objects.
        string inputPath = @"C:\Docs\OfficeMathDocument.docx";

        // Load the document. If the source contains EquationXML shapes and you want them
        // converted to OfficeMath objects, enable the flag on LoadOptions.
        LoadOptions loadOptions = new LoadOptions
        {
            ConvertShapeToOfficeMath = true // Convert EquationXML shapes to OfficeMath.
        };
        Document doc = new Document(inputPath, loadOptions);

        // -----------------------------------------------------------------
        // Example 1: Export OfficeMath as images in HTML.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlImageOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.Image
        };
        string htmlImagePath = @"C:\Docs\OfficeMath_AsImage.html";
        doc.Save(htmlImagePath, htmlImageOptions);
        Console.WriteLine("Saved HTML with OfficeMath as images to: " + htmlImagePath);

        // -----------------------------------------------------------------
        // Example 2: Export OfficeMath as MathML in HTML.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlMathMlOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML
        };
        string htmlMathMlPath = @"C:\Docs\OfficeMath_AsMathML.html";
        doc.Save(htmlMathMlPath, htmlMathMlOptions);
        Console.WriteLine("Saved HTML with OfficeMath as MathML to: " + htmlMathMlPath);

        // -----------------------------------------------------------------
        // Example 3: Export OfficeMath as plain text in HTML.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlTextOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.Text
        };
        string htmlTextPath = @"C:\Docs\OfficeMath_AsText.html";
        doc.Save(htmlTextPath, htmlTextOptions);
        Console.WriteLine("Saved HTML with OfficeMath as plain text to: " + htmlTextPath);

        // -----------------------------------------------------------------
        // Example 4: Export OfficeMath to Markdown as LaTeX.
        // -----------------------------------------------------------------
        MarkdownSaveOptions mdLatexOptions = new MarkdownSaveOptions
        {
            OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex
        };
        string markdownLatexPath = @"C:\Docs\OfficeMath_AsLatex.md";
        doc.Save(markdownLatexPath, mdLatexOptions);
        Console.WriteLine("Saved Markdown with OfficeMath as LaTeX to: " + markdownLatexPath);

        // -----------------------------------------------------------------
        // Example 5: Export OfficeMath to plain text file as LaTeX.
        // -----------------------------------------------------------------
        TxtSaveOptions txtLatexOptions = new TxtSaveOptions
        {
            OfficeMathExportMode = TxtOfficeMathExportMode.Latex
        };
        string txtLatexPath = @"C:\Docs\OfficeMath_AsLatex.txt";
        doc.Save(txtLatexPath, txtLatexOptions);
        Console.WriteLine("Saved TXT with OfficeMath as LaTeX to: " + txtLatexPath);
    }
}
