using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX that contains OfficeMath objects.
        string inputPath = @"MyDir\Office math.docx";

        // Directory where the converted files will be saved.
        string outputDir = @"ArtifactsDir\";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // ------------------------------------------------------------
        // 1. Export OfficeMath as MathML inside an HTML file.
        // ------------------------------------------------------------
        HtmlSaveOptions mathMlOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML
        };
        doc.Save(outputDir + "OfficeMath_MathML.html", mathMlOptions);

        // ------------------------------------------------------------
        // 2. Export OfficeMath as images (PNG) inside an HTML file.
        // ------------------------------------------------------------
        HtmlSaveOptions imageOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.Image
        };
        doc.Save(outputDir + "OfficeMath_Image.html", imageOptions);

        // ------------------------------------------------------------
        // 3. Export OfficeMath as plain text inside an HTML file.
        // ------------------------------------------------------------
        HtmlSaveOptions textOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.Text
        };
        doc.Save(outputDir + "OfficeMath_Text.html", textOptions);
    }
}
