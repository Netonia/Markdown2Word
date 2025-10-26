using Markdig;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace Markdown2Word.Services;

public class MarkdownToWordService
{
    private readonly MarkdownPipeline _pipeline;

    public MarkdownToWordService()
    {
        _pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();
    }

    public string RenderMarkdownToHtml(string markdownText)
    {
        if (string.IsNullOrWhiteSpace(markdownText))
            return string.Empty;

        return Markdown.ToHtml(markdownText, _pipeline);
    }

    public byte[] ConvertMarkdownToDocx(string markdownText)
    {
        var stream = new MemoryStream();
        
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Parse markdown line by line and convert to Word
            var lines = markdownText.Split('\n');
            var inCodeBlock = false;
            var codeBlockContent = new System.Text.StringBuilder();

            foreach (var line in lines)
            {
                var trimmedLine = line.TrimEnd();
                
                // Handle code blocks
                if (trimmedLine.StartsWith("```"))
                {
                    if (inCodeBlock)
                    {
                        // End of code block
                        AddCodeBlockParagraph(body, codeBlockContent.ToString());
                        codeBlockContent.Clear();
                        inCodeBlock = false;
                    }
                    else
                    {
                        // Start of code block
                        inCodeBlock = true;
                    }
                    continue;
                }

                if (inCodeBlock)
                {
                    codeBlockContent.AppendLine(trimmedLine);
                    continue;
                }

                // Handle headings
                if (trimmedLine.StartsWith("# "))
                {
                    AddHeadingParagraph(body, trimmedLine.Substring(2), 1);
                    continue;
                }
                else if (trimmedLine.StartsWith("## "))
                {
                    AddHeadingParagraph(body, trimmedLine.Substring(3), 2);
                    continue;
                }
                else if (trimmedLine.StartsWith("### "))
                {
                    AddHeadingParagraph(body, trimmedLine.Substring(4), 3);
                    continue;
                }
                else if (trimmedLine.StartsWith("#### "))
                {
                    AddHeadingParagraph(body, trimmedLine.Substring(5), 4);
                    continue;
                }

                // Handle lists
                if (trimmedLine.StartsWith("- ") || trimmedLine.StartsWith("* "))
                {
                    AddListParagraph(body, trimmedLine.Substring(2));
                    continue;
                }

                // Handle numbered lists
                if (System.Text.RegularExpressions.Regex.IsMatch(trimmedLine, @"^\d+\.\s"))
                {
                    var text = System.Text.RegularExpressions.Regex.Replace(trimmedLine, @"^\d+\.\s", "");
                    AddListParagraph(body, text);
                    continue;
                }

                // Handle empty lines
                if (string.IsNullOrWhiteSpace(trimmedLine))
                {
                    body.AppendChild(new Paragraph());
                    continue;
                }

                // Regular paragraph with inline formatting
                AddFormattedParagraph(body, trimmedLine);
            }

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    private void AddHeadingParagraph(Body body, string text, int level)
    {
        var paragraph = body.AppendChild(new Paragraph());
        var run = paragraph.AppendChild(new Run());
        run.AppendChild(new Text(text));

        var runProperties = run.InsertAt(new RunProperties(), 0);
        runProperties.AppendChild(new Bold());
        
        // Set font size based on heading level
        var fontSize = level switch
        {
            1 => "48", // 24pt
            2 => "40", // 20pt
            3 => "32", // 16pt
            4 => "28", // 14pt
            _ => "24"  // 12pt
        };
        runProperties.AppendChild(new FontSize { Val = fontSize });

        var paragraphProperties = paragraph.InsertAt(new ParagraphProperties(), 0);
        paragraphProperties.AppendChild(new SpacingBetweenLines { After = "200" });
    }

    private void AddListParagraph(Body body, string text)
    {
        var paragraph = body.AppendChild(new Paragraph());
        AddFormattedTextToRun(paragraph, text);

        var paragraphProperties = paragraph.InsertAt(new ParagraphProperties(), 0);
        paragraphProperties.AppendChild(new Indentation { Left = "720" }); // 0.5 inch
    }

    private void AddCodeBlockParagraph(Body body, string code)
    {
        var paragraph = body.AppendChild(new Paragraph());
        var run = paragraph.AppendChild(new Run());
        run.AppendChild(new Text(code) { Space = SpaceProcessingModeValues.Preserve });

        var runProperties = run.InsertAt(new RunProperties(), 0);
        runProperties.AppendChild(new RunFonts { Ascii = "Courier New" });
        runProperties.AppendChild(new FontSize { Val = "20" }); // 10pt

        var paragraphProperties = paragraph.InsertAt(new ParagraphProperties(), 0);
        paragraphProperties.AppendChild(new Indentation { Left = "720" }); // 0.5 inch
    }

    private void AddFormattedParagraph(Body body, string text)
    {
        var paragraph = body.AppendChild(new Paragraph());
        AddFormattedTextToRun(paragraph, text);
    }

    private void AddFormattedTextToRun(Paragraph paragraph, string text)
    {
        var i = 0;
        while (i < text.Length)
        {
            // Check for bold **
            if (i < text.Length - 1 && text[i] == '*' && text[i + 1] == '*')
            {
                i += 2;
                var endIndex = text.IndexOf("**", i);
                if (endIndex != -1)
                {
                    var boldText = text.Substring(i, endIndex - i);
                    var run = paragraph.AppendChild(new Run());
                    run.AppendChild(new Text(boldText));
                    var runProperties = run.InsertAt(new RunProperties(), 0);
                    runProperties.AppendChild(new Bold());
                    i = endIndex + 2;
                    continue;
                }
            }

            // Check for italic *
            if (text[i] == '*')
            {
                i++;
                var endIndex = text.IndexOf('*', i);
                if (endIndex != -1)
                {
                    var italicText = text.Substring(i, endIndex - i);
                    var run = paragraph.AppendChild(new Run());
                    run.AppendChild(new Text(italicText));
                    var runProperties = run.InsertAt(new RunProperties(), 0);
                    runProperties.AppendChild(new Italic());
                    i = endIndex + 1;
                    continue;
                }
            }

            // Check for code `
            if (text[i] == '`')
            {
                i++;
                var endIndex = text.IndexOf('`', i);
                if (endIndex != -1)
                {
                    var codeText = text.Substring(i, endIndex - i);
                    var run = paragraph.AppendChild(new Run());
                    run.AppendChild(new Text(codeText));
                    var runProperties = run.InsertAt(new RunProperties(), 0);
                    runProperties.AppendChild(new RunFonts { Ascii = "Courier New" });
                    runProperties.AppendChild(new FontSize { Val = "20" }); // 10pt
                    i = endIndex + 1;
                    continue;
                }
            }

            // Regular text
            var endPos = i + 1;
            while (endPos < text.Length && text[endPos] != '*' && text[endPos] != '`')
            {
                endPos++;
            }

            var regularText = text.Substring(i, endPos - i);
            var regularRun = paragraph.AppendChild(new Run());
            regularRun.AppendChild(new Text(regularText));
            i = endPos;
        }
    }
}
