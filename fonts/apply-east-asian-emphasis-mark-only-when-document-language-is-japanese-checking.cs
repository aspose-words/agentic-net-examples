using System.Globalization;
using Aspose.Words;
using Aspose.Words.Fonts;

// Create a new blank document.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Define the LCID for Japanese (ja-JP).
int japaneseLcid = new CultureInfo("ja-JP").LCID;

// Apply the Japanese locale to the builder's font.
// This simulates that the text we are about to insert is Japanese.
builder.Font.LocaleId = japaneseLcid;

// Apply an East Asian emphasis mark only if the current font locale is Japanese.
if (builder.Font.LocaleId == japaneseLcid)
{
    // Example emphasis mark: a solid black circle above the text.
    builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;
}

// Insert Japanese text that will display with the emphasis mark.
builder.Writeln("日本語の強調マーク");

// Clear formatting so subsequent text is not affected.
builder.Font.ClearFormatting();

// Insert normal (non‑Japanese) text for comparison.
builder.Writeln("普通のテキスト");

// Save the document.
doc.Save("EastAsianEmphasis_Japanese.docx");
