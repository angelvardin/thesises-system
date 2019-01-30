using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using DissProject.Models;
using System.IO;

namespace DocumentGeneration
{
    public class ThesisApplicationGenerator
    {
        Random random = new Random();

        public DissProject.Models.Document CreatePackage( ThesisApplication thesisApplication )
        {
            using (MemoryStream documentStream = new MemoryStream() )
            {
                using( WordprocessingDocument package = WordprocessingDocument.Create(documentStream, WordprocessingDocumentType.Document ) )
                {
                    CreateParts(package, thesisApplication );
                }

                DissProject.Models.Document result = new DissProject.Models.Document();
                result.Data = documentStream.ToArray();
                result.DateCreated = DateTime.Now;
                result.DateLastModified = DateTime.Now;
                result.Filename = "ThesisApplication_" + random.Next( 1000, 9999 ).ToString() + ".docx";
                
                return result;
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document, ThesisApplication application )
        {
            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1, application );

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("docRId0");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("docRId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            SetPackageProperties(document);
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1, ThesisApplication thesisApplication )
        {
            DocumentFormat.OpenXml.Wordprocessing.Document document1 = new DocumentFormat.OpenXml.Wordprocessing.Document();
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            WidowControl widowControl1 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens1 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation1 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification1 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color1 = new Color(){ Val = "auto" };
            Spacing spacing1 = new Spacing(){ Val = 0 };
            Position position1 = new Position(){ Val = "0" };
            FontSize fontSize1 = new FontSize(){ Val = "24" };
            Shading shading1 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(color1);
            paragraphMarkRunProperties1.Append(spacing1);
            paragraphMarkRunProperties1.Append(position1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(shading1);

            paragraphProperties1.Append(widowControl1);
            paragraphProperties1.Append(suppressAutoHyphens1);
            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color2 = new Color(){ Val = "auto" };
            Spacing spacing2 = new Spacing(){ Val = 0 };
            Position position2 = new Position(){ Val = "0" };
            FontSize fontSize2 = new FontSize(){ Val = "24" };
            Shading shading2 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(color2);
            runProperties1.Append(spacing2);
            runProperties1.Append(position2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(shading2);
            TabChar tabChar1 = new TabChar();
            TabChar tabChar2 = new TabChar();
            TabChar tabChar3 = new TabChar();

            run1.Append(runProperties1);
            run1.Append(tabChar1);
            run1.Append(tabChar2);
            run1.Append(tabChar3);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Paragraph paragraph2 = new Paragraph();

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            WidowControl widowControl2 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens2 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation2 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification2 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color3 = new Color(){ Val = "auto" };
            Spacing spacing3 = new Spacing(){ Val = 0 };
            Position position3 = new Position(){ Val = "0" };
            FontSize fontSize3 = new FontSize(){ Val = "24" };
            Shading shading3 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(color3);
            paragraphMarkRunProperties2.Append(spacing3);
            paragraphMarkRunProperties2.Append(position3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(shading3);

            paragraphProperties2.Append(widowControl2);
            paragraphProperties2.Append(suppressAutoHyphens2);
            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            paragraph2.Append(paragraphProperties2);

            Paragraph paragraph3 = new Paragraph();

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            WidowControl widowControl3 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens3 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation3 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification3 = new Justification(){ Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts4 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold1 = new Bold();
            Color color4 = new Color(){ Val = "auto" };
            Spacing spacing4 = new Spacing(){ Val = 0 };
            Position position4 = new Position(){ Val = "0" };
            FontSize fontSize4 = new FontSize(){ Val = "36" };
            Shading shading4 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties3.Append(runFonts4);
            paragraphMarkRunProperties3.Append(bold1);
            paragraphMarkRunProperties3.Append(color4);
            paragraphMarkRunProperties3.Append(spacing4);
            paragraphMarkRunProperties3.Append(position4);
            paragraphMarkRunProperties3.Append(fontSize4);
            paragraphMarkRunProperties3.Append(shading4);

            paragraphProperties3.Append(widowControl3);
            paragraphProperties3.Append(suppressAutoHyphens3);
            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts5 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold2 = new Bold();
            Color color5 = new Color(){ Val = "auto" };
            Spacing spacing5 = new Spacing(){ Val = 0 };
            Position position5 = new Position(){ Val = "0" };
            FontSize fontSize5 = new FontSize(){ Val = "36" };
            Shading shading5 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties2.Append(runFonts5);
            runProperties2.Append(bold2);
            runProperties2.Append(color5);
            runProperties2.Append(spacing5);
            runProperties2.Append(position5);
            runProperties2.Append(fontSize5);
            runProperties2.Append(shading5);
            Text text1 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text1.Text = "Предложение ";

            run2.Append(runProperties2);
            run2.Append(text1);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run2);

            Paragraph paragraph4 = new Paragraph();

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            WidowControl widowControl4 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens4 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation4 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification4 = new Justification(){ Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts6 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold3 = new Bold();
            Color color6 = new Color(){ Val = "auto" };
            Spacing spacing6 = new Spacing(){ Val = 0 };
            Position position6 = new Position(){ Val = "0" };
            FontSize fontSize6 = new FontSize(){ Val = "36" };
            Shading shading6 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties4.Append(runFonts6);
            paragraphMarkRunProperties4.Append(bold3);
            paragraphMarkRunProperties4.Append(color6);
            paragraphMarkRunProperties4.Append(spacing6);
            paragraphMarkRunProperties4.Append(position6);
            paragraphMarkRunProperties4.Append(fontSize6);
            paragraphMarkRunProperties4.Append(shading6);

            paragraphProperties4.Append(widowControl4);
            paragraphProperties4.Append(suppressAutoHyphens4);
            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(indentation4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts7 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold4 = new Bold();
            Color color7 = new Color(){ Val = "auto" };
            Spacing spacing7 = new Spacing(){ Val = 0 };
            Position position7 = new Position(){ Val = "0" };
            FontSize fontSize7 = new FontSize(){ Val = "36" };
            Shading shading7 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties3.Append(runFonts7);
            runProperties3.Append(bold4);
            runProperties3.Append(color7);
            runProperties3.Append(spacing7);
            runProperties3.Append(position7);
            runProperties3.Append(fontSize7);
            runProperties3.Append(shading7);
            Text text2 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text2.Text = "за ";

            run3.Append(runProperties3);
            run3.Append(text2);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run3);

            Paragraph paragraph5 = new Paragraph();

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            WidowControl widowControl5 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens5 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation5 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification5 = new Justification(){ Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts8 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color8 = new Color(){ Val = "auto" };
            Spacing spacing8 = new Spacing(){ Val = 0 };
            Position position8 = new Position(){ Val = "0" };
            FontSize fontSize8 = new FontSize(){ Val = "24" };
            Shading shading8 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties5.Append(runFonts8);
            paragraphMarkRunProperties5.Append(color8);
            paragraphMarkRunProperties5.Append(spacing8);
            paragraphMarkRunProperties5.Append(position8);
            paragraphMarkRunProperties5.Append(fontSize8);
            paragraphMarkRunProperties5.Append(shading8);

            paragraphProperties5.Append(widowControl5);
            paragraphProperties5.Append(suppressAutoHyphens5);
            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(indentation5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts9 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color9 = new Color(){ Val = "auto" };
            Spacing spacing9 = new Spacing(){ Val = 0 };
            Position position9 = new Position(){ Val = "0" };
            FontSize fontSize9 = new FontSize(){ Val = "24" };
            Shading shading9 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties4.Append(runFonts9);
            runProperties4.Append(color9);
            runProperties4.Append(spacing9);
            runProperties4.Append(position9);
            runProperties4.Append(fontSize9);
            runProperties4.Append(shading9);
            Text text3 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text3.Text = "възлагане на дипломна работа";

            run4.Append(runProperties4);
            run4.Append(text3);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run4);

            Paragraph paragraph6 = new Paragraph();

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            WidowControl widowControl6 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens6 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation6 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification6 = new Justification(){ Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts10 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color10 = new Color(){ Val = "auto" };
            Spacing spacing10 = new Spacing(){ Val = 0 };
            Position position10 = new Position(){ Val = "0" };
            FontSize fontSize10 = new FontSize(){ Val = "24" };
            Shading shading10 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties6.Append(runFonts10);
            paragraphMarkRunProperties6.Append(color10);
            paragraphMarkRunProperties6.Append(spacing10);
            paragraphMarkRunProperties6.Append(position10);
            paragraphMarkRunProperties6.Append(fontSize10);
            paragraphMarkRunProperties6.Append(shading10);

            paragraphProperties6.Append(widowControl6);
            paragraphProperties6.Append(suppressAutoHyphens6);
            paragraphProperties6.Append(spacingBetweenLines6);
            paragraphProperties6.Append(indentation6);
            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run5 = new Run();

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts11 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color11 = new Color(){ Val = "auto" };
            Spacing spacing11 = new Spacing(){ Val = 0 };
            Position position11 = new Position(){ Val = "0" };
            FontSize fontSize11 = new FontSize(){ Val = "24" };
            Shading shading11 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties5.Append(runFonts11);
            runProperties5.Append(color11);
            runProperties5.Append(spacing11);
            runProperties5.Append(position11);
            runProperties5.Append(fontSize11);
            runProperties5.Append(shading11);
            Text text4 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text4.Text = "към катедра “Софтуерни технологии”,";

            run5.Append(runProperties5);
            run5.Append(text4);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run5);

            Paragraph paragraph7 = new Paragraph();

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            WidowControl widowControl7 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens7 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation7 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification7 = new Justification(){ Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts12 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color12 = new Color(){ Val = "auto" };
            Spacing spacing12 = new Spacing(){ Val = 0 };
            Position position12 = new Position(){ Val = "0" };
            FontSize fontSize12 = new FontSize(){ Val = "24" };
            Shading shading12 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties7.Append(runFonts12);
            paragraphMarkRunProperties7.Append(color12);
            paragraphMarkRunProperties7.Append(spacing12);
            paragraphMarkRunProperties7.Append(position12);
            paragraphMarkRunProperties7.Append(fontSize12);
            paragraphMarkRunProperties7.Append(shading12);

            paragraphProperties7.Append(widowControl7);
            paragraphProperties7.Append(suppressAutoHyphens7);
            paragraphProperties7.Append(spacingBetweenLines7);
            paragraphProperties7.Append(indentation7);
            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts13 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color13 = new Color(){ Val = "auto" };
            Spacing spacing13 = new Spacing(){ Val = 0 };
            Position position13 = new Position(){ Val = "0" };
            FontSize fontSize13 = new FontSize(){ Val = "24" };
            Shading shading13 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties6.Append(runFonts13);
            runProperties6.Append(color13);
            runProperties6.Append(spacing13);
            runProperties6.Append(position13);
            runProperties6.Append(fontSize13);
            runProperties6.Append(shading13);
            Text text5 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text5.Text = "ФМИ, СУ “Св. Климент Охридски”";

            run6.Append(runProperties6);
            run6.Append(text5);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run6);

            Paragraph paragraph8 = new Paragraph();

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            WidowControl widowControl8 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens8 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation8 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification8 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts14 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color14 = new Color(){ Val = "auto" };
            Spacing spacing14 = new Spacing(){ Val = 0 };
            Position position14 = new Position(){ Val = "0" };
            FontSize fontSize14 = new FontSize(){ Val = "24" };
            Shading shading14 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties8.Append(runFonts14);
            paragraphMarkRunProperties8.Append(color14);
            paragraphMarkRunProperties8.Append(spacing14);
            paragraphMarkRunProperties8.Append(position14);
            paragraphMarkRunProperties8.Append(fontSize14);
            paragraphMarkRunProperties8.Append(shading14);

            paragraphProperties8.Append(widowControl8);
            paragraphProperties8.Append(suppressAutoHyphens8);
            paragraphProperties8.Append(spacingBetweenLines8);
            paragraphProperties8.Append(indentation8);
            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            paragraph8.Append(paragraphProperties8);

            Paragraph paragraph9 = new Paragraph();

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            WidowControl widowControl9 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens9 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation9 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification9 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts15 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color15 = new Color(){ Val = "auto" };
            Spacing spacing15 = new Spacing(){ Val = 0 };
            Position position15 = new Position(){ Val = "0" };
            FontSize fontSize15 = new FontSize(){ Val = "24" };
            Shading shading15 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties9.Append(runFonts15);
            paragraphMarkRunProperties9.Append(color15);
            paragraphMarkRunProperties9.Append(spacing15);
            paragraphMarkRunProperties9.Append(position15);
            paragraphMarkRunProperties9.Append(fontSize15);
            paragraphMarkRunProperties9.Append(shading15);

            paragraphProperties9.Append(widowControl9);
            paragraphProperties9.Append(suppressAutoHyphens9);
            paragraphProperties9.Append(spacingBetweenLines9);
            paragraphProperties9.Append(indentation9);
            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            paragraph9.Append(paragraphProperties9);

            Paragraph paragraph10 = new Paragraph();

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            WidowControl widowControl10 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens10 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation10 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification10 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold5 = new Bold();
            Color color16 = new Color(){ Val = "auto" };
            Spacing spacing16 = new Spacing(){ Val = 0 };
            Position position16 = new Position(){ Val = "0" };
            FontSize fontSize16 = new FontSize(){ Val = "24" };
            Shading shading16 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties10.Append(runFonts16);
            paragraphMarkRunProperties10.Append(bold5);
            paragraphMarkRunProperties10.Append(color16);
            paragraphMarkRunProperties10.Append(spacing16);
            paragraphMarkRunProperties10.Append(position16);
            paragraphMarkRunProperties10.Append(fontSize16);
            paragraphMarkRunProperties10.Append(shading16);

            paragraphProperties10.Append(widowControl10);
            paragraphProperties10.Append(suppressAutoHyphens10);
            paragraphProperties10.Append(spacingBetweenLines10);
            paragraphProperties10.Append(indentation10);
            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            paragraph10.Append(paragraphProperties10);

            Paragraph paragraph11 = new Paragraph();

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            WidowControl widowControl11 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens11 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation11 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification11 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts17 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color17 = new Color(){ Val = "auto" };
            Spacing spacing17 = new Spacing(){ Val = 0 };
            Position position17 = new Position(){ Val = "0" };
            FontSize fontSize17 = new FontSize(){ Val = "24" };
            Shading shading17 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties11.Append(runFonts17);
            paragraphMarkRunProperties11.Append(color17);
            paragraphMarkRunProperties11.Append(spacing17);
            paragraphMarkRunProperties11.Append(position17);
            paragraphMarkRunProperties11.Append(fontSize17);
            paragraphMarkRunProperties11.Append(shading17);

            paragraphProperties11.Append(widowControl11);
            paragraphProperties11.Append(suppressAutoHyphens11);
            paragraphProperties11.Append(spacingBetweenLines11);
            paragraphProperties11.Append(indentation11);
            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts18 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold6 = new Bold();
            Color color18 = new Color(){ Val = "auto" };
            Spacing spacing18 = new Spacing(){ Val = 0 };
            Position position18 = new Position(){ Val = "0" };
            FontSize fontSize18 = new FontSize(){ Val = "24" };
            Shading shading18 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties7.Append(runFonts18);
            runProperties7.Append(bold6);
            runProperties7.Append(color18);
            runProperties7.Append(spacing18);
            runProperties7.Append(position18);
            runProperties7.Append(fontSize18);
            runProperties7.Append(shading18);
            Text text6 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text6.Text = "От дипломант:";

            run7.Append(runProperties7);
            run7.Append(text6);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts19 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color19 = new Color(){ Val = "auto" };
            Spacing spacing19 = new Spacing(){ Val = 0 };
            Position position19 = new Position(){ Val = "0" };
            FontSize fontSize19 = new FontSize(){ Val = "24" };
            Shading shading19 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties8.Append(runFonts19);
            runProperties8.Append(color19);
            runProperties8.Append(spacing19);
            runProperties8.Append(position19);
            runProperties8.Append(fontSize19);
            runProperties8.Append(shading19);
            Text text7 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text7.Text = " Георги Борисов Синеклиев , специалност Софтуерно инженерство, факултетен № 61381";
            text7.Text = thesisApplication.Student.AllNames
                          + ", специалност " + thesisApplication.Student.SubjectOfStudies
                          + ", факултетен №" + thesisApplication.Student.FacultyNumber;

            run8.Append(runProperties8);
            run8.Append(text7);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run7);
            paragraph11.Append(run8);

            Paragraph paragraph12 = new Paragraph();

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            WidowControl widowControl12 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens12 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation12 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification12 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts20 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color20 = new Color(){ Val = "auto" };
            Spacing spacing20 = new Spacing(){ Val = 0 };
            Position position20 = new Position(){ Val = "0" };
            FontSize fontSize20 = new FontSize(){ Val = "24" };
            Shading shading20 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties12.Append(runFonts20);
            paragraphMarkRunProperties12.Append(color20);
            paragraphMarkRunProperties12.Append(spacing20);
            paragraphMarkRunProperties12.Append(position20);
            paragraphMarkRunProperties12.Append(fontSize20);
            paragraphMarkRunProperties12.Append(shading20);

            paragraphProperties12.Append(widowControl12);
            paragraphProperties12.Append(suppressAutoHyphens12);
            paragraphProperties12.Append(spacingBetweenLines12);
            paragraphProperties12.Append(indentation12);
            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            paragraph12.Append(paragraphProperties12);

            Paragraph paragraph13 = new Paragraph();

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            WidowControl widowControl13 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens13 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation13 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification13 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts21 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color21 = new Color(){ Val = "auto" };
            Spacing spacing21 = new Spacing(){ Val = 0 };
            Position position21 = new Position(){ Val = "0" };
            FontSize fontSize21 = new FontSize(){ Val = "24" };
            Shading shading21 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties13.Append(runFonts21);
            paragraphMarkRunProperties13.Append(color21);
            paragraphMarkRunProperties13.Append(spacing21);
            paragraphMarkRunProperties13.Append(position21);
            paragraphMarkRunProperties13.Append(fontSize21);
            paragraphMarkRunProperties13.Append(shading21);

            paragraphProperties13.Append(widowControl13);
            paragraphProperties13.Append(suppressAutoHyphens13);
            paragraphProperties13.Append(spacingBetweenLines13);
            paragraphProperties13.Append(indentation13);
            paragraphProperties13.Append(justification13);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph13.Append(paragraphProperties13);

            Paragraph paragraph14 = new Paragraph();

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            WidowControl widowControl14 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens14 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation14 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification14 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts22 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color22 = new Color(){ Val = "auto" };
            Spacing spacing22 = new Spacing(){ Val = 0 };
            Position position22 = new Position(){ Val = "0" };
            FontSize fontSize22 = new FontSize(){ Val = "24" };
            Shading shading22 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties14.Append(runFonts22);
            paragraphMarkRunProperties14.Append(color22);
            paragraphMarkRunProperties14.Append(spacing22);
            paragraphMarkRunProperties14.Append(position22);
            paragraphMarkRunProperties14.Append(fontSize22);
            paragraphMarkRunProperties14.Append(shading22);

            paragraphProperties14.Append(widowControl14);
            paragraphProperties14.Append(suppressAutoHyphens14);
            paragraphProperties14.Append(spacingBetweenLines14);
            paragraphProperties14.Append(indentation14);
            paragraphProperties14.Append(justification14);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts23 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold7 = new Bold();
            Color color23 = new Color(){ Val = "auto" };
            Spacing spacing23 = new Spacing(){ Val = 0 };
            Position position23 = new Position(){ Val = "0" };
            FontSize fontSize23 = new FontSize(){ Val = "24" };
            Shading shading23 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties9.Append(runFonts23);
            runProperties9.Append(bold7);
            runProperties9.Append(color23);
            runProperties9.Append(spacing23);
            runProperties9.Append(position23);
            runProperties9.Append(fontSize23);
            runProperties9.Append(shading23);
            Text text8 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text8.Text = "Научен ръководител:";

            run9.Append(runProperties9);
            run9.Append(text8);

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts24 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color24 = new Color(){ Val = "auto" };
            Spacing spacing24 = new Spacing(){ Val = 0 };
            Position position24 = new Position(){ Val = "0" };
            FontSize fontSize24 = new FontSize(){ Val = "24" };
            Shading shading24 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties10.Append(runFonts24);
            runProperties10.Append(color24);
            runProperties10.Append(spacing24);
            runProperties10.Append(position24);
            runProperties10.Append(fontSize24);
            runProperties10.Append(shading24);
            Text text9 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text9.Text = thesisApplication.Manager.Names + ", катедра/ВУЗ/институт";

            run10.Append(runProperties10);
            run10.Append(text9);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts25 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color25 = new Color(){ Val = "auto" };
            Spacing spacing25 = new Spacing(){ Val = 0 };
            Position position25 = new Position(){ Val = "0" };
            FontSize fontSize25 = new FontSize(){ Val = "24" };
            Shading shading25 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties11.Append(runFonts25);
            runProperties11.Append(color25);
            runProperties11.Append(spacing25);
            runProperties11.Append(position25);
            runProperties11.Append(fontSize25);
            runProperties11.Append(shading25);
            Text text10 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text10.Text = ":";

            run11.Append(runProperties11);
            run11.Append(text10);

            Run run12 = new Run();

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts26 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color26 = new Color(){ Val = "auto" };
            Spacing spacing26 = new Spacing(){ Val = 0 };
            Position position26 = new Position(){ Val = "0" };
            FontSize fontSize26 = new FontSize(){ Val = "24" };
            Shading shading26 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties12.Append(runFonts26);
            runProperties12.Append(color26);
            runProperties12.Append(spacing26);
            runProperties12.Append(position26);
            runProperties12.Append(fontSize26);
            runProperties12.Append(shading26);
            Text text11 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text11.Text = " " + thesisApplication.Manager.Department.Description;

            run12.Append(runProperties12);
            run12.Append(text11);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run9);
            paragraph14.Append(run10);
            paragraph14.Append(run11);
            paragraph14.Append(run12);

            Paragraph paragraph15 = new Paragraph();

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            WidowControl widowControl15 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens15 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation15 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification15 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color27 = new Color(){ Val = "auto" };
            Spacing spacing27 = new Spacing(){ Val = 0 };
            Position position27 = new Position(){ Val = "0" };
            FontSize fontSize27 = new FontSize(){ Val = "24" };
            Shading shading27 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties15.Append(runFonts27);
            paragraphMarkRunProperties15.Append(color27);
            paragraphMarkRunProperties15.Append(spacing27);
            paragraphMarkRunProperties15.Append(position27);
            paragraphMarkRunProperties15.Append(fontSize27);
            paragraphMarkRunProperties15.Append(shading27);

            paragraphProperties15.Append(widowControl15);
            paragraphProperties15.Append(suppressAutoHyphens15);
            paragraphProperties15.Append(spacingBetweenLines15);
            paragraphProperties15.Append(indentation15);
            paragraphProperties15.Append(justification15);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            paragraph15.Append(paragraphProperties15);

            Paragraph paragraph16 = new Paragraph();

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            WidowControl widowControl16 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens16 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation16 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification16 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold8 = new Bold();
            Color color28 = new Color(){ Val = "auto" };
            Spacing spacing28 = new Spacing(){ Val = 0 };
            Position position28 = new Position(){ Val = "0" };
            FontSize fontSize28 = new FontSize(){ Val = "24" };
            Shading shading28 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties16.Append(runFonts28);
            paragraphMarkRunProperties16.Append(bold8);
            paragraphMarkRunProperties16.Append(color28);
            paragraphMarkRunProperties16.Append(spacing28);
            paragraphMarkRunProperties16.Append(position28);
            paragraphMarkRunProperties16.Append(fontSize28);
            paragraphMarkRunProperties16.Append(shading28);

            paragraphProperties16.Append(widowControl16);
            paragraphProperties16.Append(suppressAutoHyphens16);
            paragraphProperties16.Append(spacingBetweenLines16);
            paragraphProperties16.Append(indentation16);
            paragraphProperties16.Append(justification16);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            paragraph16.Append(paragraphProperties16);

            Paragraph paragraph17 = new Paragraph();

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            WidowControl widowControl17 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens17 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation17 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification17 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts29 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color29 = new Color(){ Val = "auto" };
            Spacing spacing29 = new Spacing(){ Val = 0 };
            Position position29 = new Position(){ Val = "0" };
            FontSize fontSize29 = new FontSize(){ Val = "24" };
            Shading shading29 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties17.Append(runFonts29);
            paragraphMarkRunProperties17.Append(color29);
            paragraphMarkRunProperties17.Append(spacing29);
            paragraphMarkRunProperties17.Append(position29);
            paragraphMarkRunProperties17.Append(fontSize29);
            paragraphMarkRunProperties17.Append(shading29);

            paragraphProperties17.Append(widowControl17);
            paragraphProperties17.Append(suppressAutoHyphens17);
            paragraphProperties17.Append(spacingBetweenLines17);
            paragraphProperties17.Append(indentation17);
            paragraphProperties17.Append(justification17);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts30 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold9 = new Bold();
            Color color30 = new Color(){ Val = "auto" };
            Spacing spacing30 = new Spacing(){ Val = 0 };
            Position position30 = new Position(){ Val = "0" };
            FontSize fontSize30 = new FontSize(){ Val = "24" };
            Shading shading30 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties13.Append(runFonts30);
            runProperties13.Append(bold9);
            runProperties13.Append(color30);
            runProperties13.Append(spacing30);
            runProperties13.Append(position30);
            runProperties13.Append(fontSize30);
            runProperties13.Append(shading30);
            Text text12 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text12.Text = "Консултант:  ";

            run13.Append(runProperties13);
            run13.Append(text12);

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts31 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color31 = new Color(){ Val = "auto" };
            Spacing spacing31 = new Spacing(){ Val = 0 };
            Position position31 = new Position(){ Val = "0" };
            FontSize fontSize31 = new FontSize(){ Val = "24" };
            Shading shading31 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties14.Append(runFonts31);
            runProperties14.Append(color31);
            runProperties14.Append(spacing31);
            runProperties14.Append(position31);
            runProperties14.Append(fontSize31);
            runProperties14.Append(shading31);
            Text text13 = new Text(){ Space = SpaceProcessingModeValues.Preserve };

            string consultantNames = "";
            string consultantDepartment = "";
            if (thesisApplication.Consultants != null)
            {
                Person consultant = thesisApplication.Consultants.First();
                if (consultant != null)
                {
                    consultantNames = consultant.Names;
                }

                try
                {
                    Teacher consultantTeacher = (Teacher)consultant;
                    if (consultantTeacher != null)
                    {
                        consultantDepartment = consultantTeacher.Department.Description;
                    }
                }
                catch (Exception ex)
                {
                    // do nothing
                }
            }

            text13.Text = consultantNames + " , катедра/ВУЗ/институт";

            run14.Append(runProperties14);
            run14.Append(text13);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts32 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color32 = new Color(){ Val = "auto" };
            Spacing spacing32 = new Spacing(){ Val = 0 };
            Position position32 = new Position(){ Val = "0" };
            FontSize fontSize32 = new FontSize(){ Val = "24" };
            Shading shading32 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties15.Append(runFonts32);
            runProperties15.Append(color32);
            runProperties15.Append(spacing32);
            runProperties15.Append(position32);
            runProperties15.Append(fontSize32);
            runProperties15.Append(shading32);
            Text text14 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text14.Text = ":";

            run15.Append(runProperties15);
            run15.Append(text14);

            Run run16 = new Run();

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts33 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color33 = new Color(){ Val = "auto" };
            Spacing spacing33 = new Spacing(){ Val = 0 };
            Position position33 = new Position(){ Val = "0" };
            FontSize fontSize33 = new FontSize(){ Val = "24" };
            Shading shading33 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties16.Append(runFonts33);
            runProperties16.Append(color33);
            runProperties16.Append(spacing33);
            runProperties16.Append(position33);
            runProperties16.Append(fontSize33);
            runProperties16.Append(shading33);
            Text text15 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text15.Text = " " + consultantDepartment;

            run16.Append(runProperties16);
            run16.Append(text15);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run13);
            paragraph17.Append(run14);
            paragraph17.Append(run15);
            paragraph17.Append(run16);

            Paragraph paragraph18 = new Paragraph();

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            WidowControl widowControl18 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens18 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation18 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification18 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts34 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color34 = new Color(){ Val = "auto" };
            Spacing spacing34 = new Spacing(){ Val = 0 };
            Position position34 = new Position(){ Val = "0" };
            FontSize fontSize34 = new FontSize(){ Val = "24" };
            Shading shading34 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties18.Append(runFonts34);
            paragraphMarkRunProperties18.Append(color34);
            paragraphMarkRunProperties18.Append(spacing34);
            paragraphMarkRunProperties18.Append(position34);
            paragraphMarkRunProperties18.Append(fontSize34);
            paragraphMarkRunProperties18.Append(shading34);

            paragraphProperties18.Append(widowControl18);
            paragraphProperties18.Append(suppressAutoHyphens18);
            paragraphProperties18.Append(spacingBetweenLines18);
            paragraphProperties18.Append(indentation18);
            paragraphProperties18.Append(justification18);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            paragraph18.Append(paragraphProperties18);

            Paragraph paragraph19 = new Paragraph();

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            WidowControl widowControl19 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens19 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation19 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification19 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts35 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color35 = new Color(){ Val = "auto" };
            Spacing spacing35 = new Spacing(){ Val = 0 };
            Position position35 = new Position(){ Val = "0" };
            FontSize fontSize35 = new FontSize(){ Val = "24" };
            Shading shading35 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties19.Append(runFonts35);
            paragraphMarkRunProperties19.Append(color35);
            paragraphMarkRunProperties19.Append(spacing35);
            paragraphMarkRunProperties19.Append(position35);
            paragraphMarkRunProperties19.Append(fontSize35);
            paragraphMarkRunProperties19.Append(shading35);

            paragraphProperties19.Append(widowControl19);
            paragraphProperties19.Append(suppressAutoHyphens19);
            paragraphProperties19.Append(spacingBetweenLines19);
            paragraphProperties19.Append(indentation19);
            paragraphProperties19.Append(justification19);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            paragraph19.Append(paragraphProperties19);

            Paragraph paragraph20 = new Paragraph();

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            WidowControl widowControl20 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens20 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation20 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification20 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold10 = new Bold();
            Color color36 = new Color(){ Val = "auto" };
            Spacing spacing36 = new Spacing(){ Val = 0 };
            Position position36 = new Position(){ Val = "0" };
            FontSize fontSize36 = new FontSize(){ Val = "24" };
            Shading shading36 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties20.Append(runFonts36);
            paragraphMarkRunProperties20.Append(bold10);
            paragraphMarkRunProperties20.Append(color36);
            paragraphMarkRunProperties20.Append(spacing36);
            paragraphMarkRunProperties20.Append(position36);
            paragraphMarkRunProperties20.Append(fontSize36);
            paragraphMarkRunProperties20.Append(shading36);

            paragraphProperties20.Append(widowControl20);
            paragraphProperties20.Append(suppressAutoHyphens20);
            paragraphProperties20.Append(spacingBetweenLines20);
            paragraphProperties20.Append(indentation20);
            paragraphProperties20.Append(justification20);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run17 = new Run();

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts37 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold11 = new Bold();
            Color color37 = new Color(){ Val = "auto" };
            Spacing spacing37 = new Spacing(){ Val = 0 };
            Position position37 = new Position(){ Val = "0" };
            FontSize fontSize37 = new FontSize(){ Val = "24" };
            Shading shading37 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties17.Append(runFonts37);
            runProperties17.Append(bold11);
            runProperties17.Append(color37);
            runProperties17.Append(spacing37);
            runProperties17.Append(position37);
            runProperties17.Append(fontSize37);
            runProperties17.Append(shading37);
            Text text16 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text16.Text = "Тема на дипломната работа:";

            run17.Append(runProperties17);
            run17.Append(text16);

            Run run18 = new Run();

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts38 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold12 = new Bold();
            Color color38 = new Color(){ Val = "auto" };
            Spacing spacing38 = new Spacing(){ Val = 0 };
            Position position38 = new Position(){ Val = "0" };
            FontSize fontSize38 = new FontSize(){ Val = "24" };
            Shading shading38 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties18.Append(runFonts38);
            runProperties18.Append(bold12);
            runProperties18.Append(color38);
            runProperties18.Append(spacing38);
            runProperties18.Append(position38);
            runProperties18.Append(fontSize38);
            runProperties18.Append(shading38);
            Text text17 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text17.Text = " ";

            run18.Append(runProperties18);
            run18.Append(text17);

            Run run19 = new Run();

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts39 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold13 = new Bold();
            Color color39 = new Color(){ Val = "auto" };
            Spacing spacing39 = new Spacing(){ Val = 0 };
            Position position39 = new Position(){ Val = "0" };
            FontSize fontSize39 = new FontSize(){ Val = "24" };
            Shading shading39 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties19.Append(runFonts39);
            runProperties19.Append(bold13);
            runProperties19.Append(color39);
            runProperties19.Append(spacing39);
            runProperties19.Append(position39);
            runProperties19.Append(fontSize39);
            runProperties19.Append(shading39);
            Text text18 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text18.Text = thesisApplication.Subject;

            run19.Append(runProperties19);
            run19.Append(text18);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run17);
            paragraph20.Append(run18);
            paragraph20.Append(run19);

            Paragraph paragraph21 = new Paragraph();

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            WidowControl widowControl21 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens21 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation21 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification21 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts40 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color40 = new Color(){ Val = "auto" };
            Spacing spacing40 = new Spacing(){ Val = 0 };
            Position position40 = new Position(){ Val = "0" };
            FontSize fontSize40 = new FontSize(){ Val = "24" };
            Shading shading40 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties21.Append(runFonts40);
            paragraphMarkRunProperties21.Append(color40);
            paragraphMarkRunProperties21.Append(spacing40);
            paragraphMarkRunProperties21.Append(position40);
            paragraphMarkRunProperties21.Append(fontSize40);
            paragraphMarkRunProperties21.Append(shading40);

            paragraphProperties21.Append(widowControl21);
            paragraphProperties21.Append(suppressAutoHyphens21);
            paragraphProperties21.Append(spacingBetweenLines21);
            paragraphProperties21.Append(indentation21);
            paragraphProperties21.Append(justification21);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            paragraph21.Append(paragraphProperties21);

            Paragraph paragraph22 = new Paragraph();

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            WidowControl widowControl22 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens22 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation22 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification22 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts41 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color41 = new Color(){ Val = "auto" };
            Spacing spacing41 = new Spacing(){ Val = 0 };
            Position position41 = new Position(){ Val = "0" };
            FontSize fontSize41 = new FontSize(){ Val = "24" };
            Shading shading41 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties22.Append(runFonts41);
            paragraphMarkRunProperties22.Append(color41);
            paragraphMarkRunProperties22.Append(spacing41);
            paragraphMarkRunProperties22.Append(position41);
            paragraphMarkRunProperties22.Append(fontSize41);
            paragraphMarkRunProperties22.Append(shading41);

            paragraphProperties22.Append(widowControl22);
            paragraphProperties22.Append(suppressAutoHyphens22);
            paragraphProperties22.Append(spacingBetweenLines22);
            paragraphProperties22.Append(indentation22);
            paragraphProperties22.Append(justification22);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            paragraph22.Append(paragraphProperties22);

            Paragraph paragraph23 = new Paragraph();

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            WidowControl widowControl23 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens23 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation23 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification23 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts42 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold14 = new Bold();
            Color color42 = new Color(){ Val = "auto" };
            Spacing spacing42 = new Spacing(){ Val = 0 };
            Position position42 = new Position(){ Val = "0" };
            FontSize fontSize42 = new FontSize(){ Val = "24" };
            Shading shading42 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties23.Append(runFonts42);
            paragraphMarkRunProperties23.Append(bold14);
            paragraphMarkRunProperties23.Append(color42);
            paragraphMarkRunProperties23.Append(spacing42);
            paragraphMarkRunProperties23.Append(position42);
            paragraphMarkRunProperties23.Append(fontSize42);
            paragraphMarkRunProperties23.Append(shading42);

            paragraphProperties23.Append(widowControl23);
            paragraphProperties23.Append(suppressAutoHyphens23);
            paragraphProperties23.Append(spacingBetweenLines23);
            paragraphProperties23.Append(indentation23);
            paragraphProperties23.Append(justification23);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run20 = new Run();

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts43 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold15 = new Bold();
            Color color43 = new Color(){ Val = "auto" };
            Spacing spacing43 = new Spacing(){ Val = 0 };
            Position position43 = new Position(){ Val = "0" };
            FontSize fontSize43 = new FontSize(){ Val = "24" };
            Shading shading43 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties20.Append(runFonts43);
            runProperties20.Append(bold15);
            runProperties20.Append(color43);
            runProperties20.Append(spacing43);
            runProperties20.Append(position43);
            runProperties20.Append(fontSize43);
            runProperties20.Append(shading43);
            Text text19 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text19.Text = "Анотация: " +  thesisApplication.Annotation;

            run20.Append(runProperties20);
            run20.Append(text19);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run20);

            Paragraph paragraph24 = new Paragraph();

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            WidowControl widowControl24 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens24 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines24 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation24 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification24 = new Justification(){ Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color44 = new Color(){ Val = "auto" };
            Spacing spacing44 = new Spacing(){ Val = 0 };
            Position position44 = new Position(){ Val = "0" };
            FontSize fontSize44 = new FontSize(){ Val = "24" };
            Shading shading44 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties24.Append(runFonts44);
            paragraphMarkRunProperties24.Append(color44);
            paragraphMarkRunProperties24.Append(spacing44);
            paragraphMarkRunProperties24.Append(position44);
            paragraphMarkRunProperties24.Append(fontSize44);
            paragraphMarkRunProperties24.Append(shading44);

            paragraphProperties24.Append(widowControl24);
            paragraphProperties24.Append(suppressAutoHyphens24);
            paragraphProperties24.Append(spacingBetweenLines24);
            paragraphProperties24.Append(indentation24);
            paragraphProperties24.Append(justification24);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            paragraph24.Append(paragraphProperties24);

            Paragraph paragraph25 = new Paragraph();

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            WidowControl widowControl25 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens25 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation25 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification25 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts45 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color45 = new Color(){ Val = "auto" };
            Spacing spacing45 = new Spacing(){ Val = 0 };
            Position position45 = new Position(){ Val = "0" };
            FontSize fontSize45 = new FontSize(){ Val = "24" };
            Shading shading45 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties25.Append(runFonts45);
            paragraphMarkRunProperties25.Append(color45);
            paragraphMarkRunProperties25.Append(spacing45);
            paragraphMarkRunProperties25.Append(position45);
            paragraphMarkRunProperties25.Append(fontSize45);
            paragraphMarkRunProperties25.Append(shading45);

            paragraphProperties25.Append(widowControl25);
            paragraphProperties25.Append(suppressAutoHyphens25);
            paragraphProperties25.Append(spacingBetweenLines25);
            paragraphProperties25.Append(indentation25);
            paragraphProperties25.Append(justification25);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            paragraph25.Append(paragraphProperties25);

            Paragraph paragraph26 = new Paragraph();

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            WidowControl widowControl26 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens26 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation26 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification26 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts46 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold16 = new Bold();
            Color color46 = new Color(){ Val = "auto" };
            Spacing spacing46 = new Spacing(){ Val = 0 };
            Position position46 = new Position(){ Val = "0" };
            FontSize fontSize46 = new FontSize(){ Val = "24" };
            Shading shading46 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties26.Append(runFonts46);
            paragraphMarkRunProperties26.Append(bold16);
            paragraphMarkRunProperties26.Append(color46);
            paragraphMarkRunProperties26.Append(spacing46);
            paragraphMarkRunProperties26.Append(position46);
            paragraphMarkRunProperties26.Append(fontSize46);
            paragraphMarkRunProperties26.Append(shading46);

            paragraphProperties26.Append(widowControl26);
            paragraphProperties26.Append(suppressAutoHyphens26);
            paragraphProperties26.Append(spacingBetweenLines26);
            paragraphProperties26.Append(indentation26);
            paragraphProperties26.Append(justification26);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            Run run21 = new Run();

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts47 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold17 = new Bold();
            Color color47 = new Color(){ Val = "auto" };
            Spacing spacing47 = new Spacing(){ Val = 0 };
            Position position47 = new Position(){ Val = "0" };
            FontSize fontSize47 = new FontSize(){ Val = "24" };
            Shading shading47 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties21.Append(runFonts47);
            runProperties21.Append(bold17);
            runProperties21.Append(color47);
            runProperties21.Append(spacing47);
            runProperties21.Append(position47);
            runProperties21.Append(fontSize47);
            runProperties21.Append(shading47);
            Text text20 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text20.Text = "Цел на дипломната работа: " + thesisApplication.Purpose;

            run21.Append(runProperties21);
            run21.Append(text20);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(run21);

            Paragraph paragraph27 = new Paragraph();

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            WidowControl widowControl27 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens27 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines27 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation27 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification27 = new Justification(){ Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts48 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color48 = new Color(){ Val = "auto" };
            Spacing spacing48 = new Spacing(){ Val = 0 };
            Position position48 = new Position(){ Val = "0" };
            FontSize fontSize48 = new FontSize(){ Val = "24" };
            Shading shading48 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties27.Append(runFonts48);
            paragraphMarkRunProperties27.Append(color48);
            paragraphMarkRunProperties27.Append(spacing48);
            paragraphMarkRunProperties27.Append(position48);
            paragraphMarkRunProperties27.Append(fontSize48);
            paragraphMarkRunProperties27.Append(shading48);

            paragraphProperties27.Append(widowControl27);
            paragraphProperties27.Append(suppressAutoHyphens27);
            paragraphProperties27.Append(spacingBetweenLines27);
            paragraphProperties27.Append(indentation27);
            paragraphProperties27.Append(justification27);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            paragraph27.Append(paragraphProperties27);

            Paragraph paragraph28 = new Paragraph();

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            WidowControl widowControl28 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens28 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines28 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation28 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification28 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts49 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold18 = new Bold();
            Color color49 = new Color(){ Val = "auto" };
            Spacing spacing49 = new Spacing(){ Val = 0 };
            Position position49 = new Position(){ Val = "0" };
            FontSize fontSize49 = new FontSize(){ Val = "24" };
            Shading shading49 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties28.Append(runFonts49);
            paragraphMarkRunProperties28.Append(bold18);
            paragraphMarkRunProperties28.Append(color49);
            paragraphMarkRunProperties28.Append(spacing49);
            paragraphMarkRunProperties28.Append(position49);
            paragraphMarkRunProperties28.Append(fontSize49);
            paragraphMarkRunProperties28.Append(shading49);

            paragraphProperties28.Append(widowControl28);
            paragraphProperties28.Append(suppressAutoHyphens28);
            paragraphProperties28.Append(spacingBetweenLines28);
            paragraphProperties28.Append(indentation28);
            paragraphProperties28.Append(justification28);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            paragraph28.Append(paragraphProperties28);

            Paragraph paragraph29 = new Paragraph();

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            WidowControl widowControl29 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens29 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines29 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation29 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification29 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts50 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color50 = new Color(){ Val = "auto" };
            Spacing spacing50 = new Spacing(){ Val = 0 };
            Position position50 = new Position(){ Val = "0" };
            FontSize fontSize50 = new FontSize(){ Val = "24" };
            Shading shading50 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties29.Append(runFonts50);
            paragraphMarkRunProperties29.Append(color50);
            paragraphMarkRunProperties29.Append(spacing50);
            paragraphMarkRunProperties29.Append(position50);
            paragraphMarkRunProperties29.Append(fontSize50);
            paragraphMarkRunProperties29.Append(shading50);

            paragraphProperties29.Append(widowControl29);
            paragraphProperties29.Append(suppressAutoHyphens29);
            paragraphProperties29.Append(spacingBetweenLines29);
            paragraphProperties29.Append(indentation29);
            paragraphProperties29.Append(justification29);
            paragraphProperties29.Append(paragraphMarkRunProperties29);

            Run run22 = new Run();

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts51 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold19 = new Bold();
            Color color51 = new Color(){ Val = "auto" };
            Spacing spacing51 = new Spacing(){ Val = 0 };
            Position position51 = new Position(){ Val = "0" };
            FontSize fontSize51 = new FontSize(){ Val = "24" };
            Shading shading51 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties22.Append(runFonts51);
            runProperties22.Append(bold19);
            runProperties22.Append(color51);
            runProperties22.Append(spacing51);
            runProperties22.Append(position51);
            runProperties22.Append(fontSize51);
            runProperties22.Append(shading51);
            Text text21 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text21.Text = "Задачи, произтичащи от целта: " + thesisApplication.Tasks;

            run22.Append(runProperties22);
            run22.Append(text21);

            paragraph29.Append(paragraphProperties29);
            paragraph29.Append(run22);

            Paragraph paragraph30 = new Paragraph();

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            WidowControl widowControl30 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens30 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines30 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation30 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification30 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts52 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color52 = new Color(){ Val = "auto" };
            Spacing spacing52 = new Spacing(){ Val = 0 };
            Position position52 = new Position(){ Val = "0" };
            FontSize fontSize52 = new FontSize(){ Val = "24" };
            Shading shading52 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties30.Append(runFonts52);
            paragraphMarkRunProperties30.Append(color52);
            paragraphMarkRunProperties30.Append(spacing52);
            paragraphMarkRunProperties30.Append(position52);
            paragraphMarkRunProperties30.Append(fontSize52);
            paragraphMarkRunProperties30.Append(shading52);

            paragraphProperties30.Append(widowControl30);
            paragraphProperties30.Append(suppressAutoHyphens30);
            paragraphProperties30.Append(spacingBetweenLines30);
            paragraphProperties30.Append(indentation30);
            paragraphProperties30.Append(justification30);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            paragraph30.Append(paragraphProperties30);

            Paragraph paragraph31 = new Paragraph();

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            WidowControl widowControl31 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens31 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines31 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation31 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification31 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts53 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold20 = new Bold();
            Color color53 = new Color(){ Val = "auto" };
            Spacing spacing53 = new Spacing(){ Val = 0 };
            Position position53 = new Position(){ Val = "0" };
            FontSize fontSize53 = new FontSize(){ Val = "24" };
            Shading shading53 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties31.Append(runFonts53);
            paragraphMarkRunProperties31.Append(bold20);
            paragraphMarkRunProperties31.Append(color53);
            paragraphMarkRunProperties31.Append(spacing53);
            paragraphMarkRunProperties31.Append(position53);
            paragraphMarkRunProperties31.Append(fontSize53);
            paragraphMarkRunProperties31.Append(shading53);

            paragraphProperties31.Append(widowControl31);
            paragraphProperties31.Append(suppressAutoHyphens31);
            paragraphProperties31.Append(spacingBetweenLines31);
            paragraphProperties31.Append(indentation31);
            paragraphProperties31.Append(justification31);
            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run23 = new Run();

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts54 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold21 = new Bold();
            Color color54 = new Color(){ Val = "auto" };
            Spacing spacing54 = new Spacing(){ Val = 0 };
            Position position54 = new Position(){ Val = "0" };
            FontSize fontSize54 = new FontSize(){ Val = "24" };
            Shading shading54 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties23.Append(runFonts54);
            runProperties23.Append(bold21);
            runProperties23.Append(color54);
            runProperties23.Append(spacing54);
            runProperties23.Append(position54);
            runProperties23.Append(fontSize54);
            runProperties23.Append(shading54);
            Text text22 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text22.Text = "Ограничаващи/облекчаващи условия: " + thesisApplication.Constraints;

            run23.Append(runProperties23);
            run23.Append(text22);

            paragraph31.Append(paragraphProperties31);
            paragraph31.Append(run23);

            Paragraph paragraph32 = new Paragraph();

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            WidowControl widowControl32 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens32 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines32 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation32 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification32 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts55 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color55 = new Color(){ Val = "auto" };
            Spacing spacing55 = new Spacing(){ Val = 0 };
            Position position55 = new Position(){ Val = "0" };
            FontSize fontSize55 = new FontSize(){ Val = "24" };
            Shading shading55 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties32.Append(runFonts55);
            paragraphMarkRunProperties32.Append(color55);
            paragraphMarkRunProperties32.Append(spacing55);
            paragraphMarkRunProperties32.Append(position55);
            paragraphMarkRunProperties32.Append(fontSize55);
            paragraphMarkRunProperties32.Append(shading55);

            paragraphProperties32.Append(widowControl32);
            paragraphProperties32.Append(suppressAutoHyphens32);
            paragraphProperties32.Append(spacingBetweenLines32);
            paragraphProperties32.Append(indentation32);
            paragraphProperties32.Append(justification32);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            paragraph32.Append(paragraphProperties32);

            Paragraph paragraph33 = new Paragraph();

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            WidowControl widowControl33 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens33 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines33 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation33 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification33 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts56 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color56 = new Color(){ Val = "auto" };
            Spacing spacing56 = new Spacing(){ Val = 0 };
            Position position56 = new Position(){ Val = "0" };
            FontSize fontSize56 = new FontSize(){ Val = "24" };
            Shading shading56 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties33.Append(runFonts56);
            paragraphMarkRunProperties33.Append(color56);
            paragraphMarkRunProperties33.Append(spacing56);
            paragraphMarkRunProperties33.Append(position56);
            paragraphMarkRunProperties33.Append(fontSize56);
            paragraphMarkRunProperties33.Append(shading56);

            paragraphProperties33.Append(widowControl33);
            paragraphProperties33.Append(suppressAutoHyphens33);
            paragraphProperties33.Append(spacingBetweenLines33);
            paragraphProperties33.Append(indentation33);
            paragraphProperties33.Append(justification33);
            paragraphProperties33.Append(paragraphMarkRunProperties33);

            paragraph33.Append(paragraphProperties33);

            Paragraph paragraph34 = new Paragraph();

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            WidowControl widowControl34 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens34 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines34 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation34 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification34 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts57 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold22 = new Bold();
            Color color57 = new Color(){ Val = "auto" };
            Spacing spacing57 = new Spacing(){ Val = 0 };
            Position position57 = new Position(){ Val = "0" };
            FontSize fontSize57 = new FontSize(){ Val = "24" };
            Shading shading57 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties34.Append(runFonts57);
            paragraphMarkRunProperties34.Append(bold22);
            paragraphMarkRunProperties34.Append(color57);
            paragraphMarkRunProperties34.Append(spacing57);
            paragraphMarkRunProperties34.Append(position57);
            paragraphMarkRunProperties34.Append(fontSize57);
            paragraphMarkRunProperties34.Append(shading57);

            paragraphProperties34.Append(widowControl34);
            paragraphProperties34.Append(suppressAutoHyphens34);
            paragraphProperties34.Append(spacingBetweenLines34);
            paragraphProperties34.Append(indentation34);
            paragraphProperties34.Append(justification34);
            paragraphProperties34.Append(paragraphMarkRunProperties34);

            Run run24 = new Run();

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts58 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold23 = new Bold();
            Color color58 = new Color(){ Val = "auto" };
            Spacing spacing58 = new Spacing(){ Val = 0 };
            Position position58 = new Position(){ Val = "0" };
            FontSize fontSize58 = new FontSize(){ Val = "24" };
            Shading shading58 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties24.Append(runFonts58);
            runProperties24.Append(bold23);
            runProperties24.Append(color58);
            runProperties24.Append(spacing58);
            runProperties24.Append(position58);
            runProperties24.Append(fontSize58);
            runProperties24.Append(shading58);
            Text text23 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text23.Text = "Срок за изпълнение: " + thesisApplication.Deadline.ToString( "dd.MM.yyyy" );

            run24.Append(runProperties24);
            run24.Append(text23);

            paragraph34.Append(paragraphProperties34);
            paragraph34.Append(run24);

            Paragraph paragraph35 = new Paragraph();

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            WidowControl widowControl35 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens35 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines35 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation35 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification35 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts59 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color59 = new Color(){ Val = "auto" };
            Spacing spacing59 = new Spacing(){ Val = 0 };
            Position position59 = new Position(){ Val = "0" };
            FontSize fontSize59 = new FontSize(){ Val = "24" };
            Shading shading59 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties35.Append(runFonts59);
            paragraphMarkRunProperties35.Append(color59);
            paragraphMarkRunProperties35.Append(spacing59);
            paragraphMarkRunProperties35.Append(position59);
            paragraphMarkRunProperties35.Append(fontSize59);
            paragraphMarkRunProperties35.Append(shading59);

            paragraphProperties35.Append(widowControl35);
            paragraphProperties35.Append(suppressAutoHyphens35);
            paragraphProperties35.Append(spacingBetweenLines35);
            paragraphProperties35.Append(indentation35);
            paragraphProperties35.Append(justification35);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            paragraph35.Append(paragraphProperties35);

            Paragraph paragraph36 = new Paragraph();

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            WidowControl widowControl36 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens36 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines36 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation36 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification36 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts60 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color60 = new Color(){ Val = "auto" };
            Spacing spacing60 = new Spacing(){ Val = 0 };
            Position position60 = new Position(){ Val = "0" };
            FontSize fontSize60 = new FontSize(){ Val = "24" };
            Shading shading60 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties36.Append(runFonts60);
            paragraphMarkRunProperties36.Append(color60);
            paragraphMarkRunProperties36.Append(spacing60);
            paragraphMarkRunProperties36.Append(position60);
            paragraphMarkRunProperties36.Append(fontSize60);
            paragraphMarkRunProperties36.Append(shading60);

            paragraphProperties36.Append(widowControl36);
            paragraphProperties36.Append(suppressAutoHyphens36);
            paragraphProperties36.Append(spacingBetweenLines36);
            paragraphProperties36.Append(indentation36);
            paragraphProperties36.Append(justification36);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            paragraph36.Append(paragraphProperties36);

            Paragraph paragraph37 = new Paragraph();

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            WidowControl widowControl37 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens37 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines37 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation37 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification37 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts61 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color61 = new Color(){ Val = "auto" };
            Spacing spacing61 = new Spacing(){ Val = 0 };
            Position position61 = new Position(){ Val = "0" };
            FontSize fontSize61 = new FontSize(){ Val = "24" };
            Shading shading61 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties37.Append(runFonts61);
            paragraphMarkRunProperties37.Append(color61);
            paragraphMarkRunProperties37.Append(spacing61);
            paragraphMarkRunProperties37.Append(position61);
            paragraphMarkRunProperties37.Append(fontSize61);
            paragraphMarkRunProperties37.Append(shading61);

            paragraphProperties37.Append(widowControl37);
            paragraphProperties37.Append(suppressAutoHyphens37);
            paragraphProperties37.Append(spacingBetweenLines37);
            paragraphProperties37.Append(indentation37);
            paragraphProperties37.Append(justification37);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run25 = new Run();

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts62 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color62 = new Color(){ Val = "auto" };
            Spacing spacing62 = new Spacing(){ Val = 0 };
            Position position62 = new Position(){ Val = "0" };
            FontSize fontSize62 = new FontSize(){ Val = "24" };
            Shading shading62 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties25.Append(runFonts62);
            runProperties25.Append(color62);
            runProperties25.Append(spacing62);
            runProperties25.Append(position62);
            runProperties25.Append(fontSize62);
            runProperties25.Append(shading62);
            Text text24 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text24.Text = "Дата: " + DateTime.Now.Date.ToString("dd.MM.yyyy");
            TabChar tabChar4 = new TabChar();
            TabChar tabChar5 = new TabChar();
            TabChar tabChar6 = new TabChar();
            TabChar tabChar7 = new TabChar();
            TabChar tabChar8 = new TabChar();
            TabChar tabChar9 = new TabChar();
            TabChar tabChar10 = new TabChar();
            TabChar tabChar11 = new TabChar();
            Text text25 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text25.Text = "Заявител:";

            run25.Append(runProperties25);
            run25.Append(text24);
            run25.Append(tabChar4);
            run25.Append(tabChar5);
            run25.Append(tabChar6);
            run25.Append(tabChar7);
            run25.Append(tabChar8);
            run25.Append(tabChar9);
            run25.Append(tabChar10);
            run25.Append(tabChar11);
            run25.Append(text25);

            paragraph37.Append(paragraphProperties37);
            paragraph37.Append(run25);

            Paragraph paragraph38 = new Paragraph();

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            WidowControl widowControl38 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens38 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines38 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation38 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification38 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color63 = new Color(){ Val = "auto" };
            Spacing spacing63 = new Spacing(){ Val = 0 };
            Position position63 = new Position(){ Val = "0" };
            FontSize fontSize63 = new FontSize(){ Val = "24" };
            Shading shading63 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties38.Append(runFonts63);
            paragraphMarkRunProperties38.Append(color63);
            paragraphMarkRunProperties38.Append(spacing63);
            paragraphMarkRunProperties38.Append(position63);
            paragraphMarkRunProperties38.Append(fontSize63);
            paragraphMarkRunProperties38.Append(shading63);

            paragraphProperties38.Append(widowControl38);
            paragraphProperties38.Append(suppressAutoHyphens38);
            paragraphProperties38.Append(spacingBetweenLines38);
            paragraphProperties38.Append(indentation38);
            paragraphProperties38.Append(justification38);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph38.Append(paragraphProperties38);

            Paragraph paragraph39 = new Paragraph();

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            WidowControl widowControl39 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens39 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines39 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation39 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification39 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts64 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color64 = new Color(){ Val = "auto" };
            Spacing spacing64 = new Spacing(){ Val = 0 };
            Position position64 = new Position(){ Val = "0" };
            FontSize fontSize64 = new FontSize(){ Val = "24" };
            Shading shading64 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties39.Append(runFonts64);
            paragraphMarkRunProperties39.Append(color64);
            paragraphMarkRunProperties39.Append(spacing64);
            paragraphMarkRunProperties39.Append(position64);
            paragraphMarkRunProperties39.Append(fontSize64);
            paragraphMarkRunProperties39.Append(shading64);

            paragraphProperties39.Append(widowControl39);
            paragraphProperties39.Append(suppressAutoHyphens39);
            paragraphProperties39.Append(spacingBetweenLines39);
            paragraphProperties39.Append(indentation39);
            paragraphProperties39.Append(justification39);
            paragraphProperties39.Append(paragraphMarkRunProperties39);

            Run run26 = new Run();

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts65 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color65 = new Color(){ Val = "auto" };
            Spacing spacing65 = new Spacing(){ Val = 0 };
            Position position65 = new Position(){ Val = "0" };
            FontSize fontSize65 = new FontSize(){ Val = "24" };
            Shading shading65 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties26.Append(runFonts65);
            runProperties26.Append(color65);
            runProperties26.Append(spacing65);
            runProperties26.Append(position65);
            runProperties26.Append(fontSize65);
            runProperties26.Append(shading65);
            TabChar tabChar12 = new TabChar();
            TabChar tabChar13 = new TabChar();
            TabChar tabChar14 = new TabChar();
            TabChar tabChar15 = new TabChar();
            TabChar tabChar16 = new TabChar();
            TabChar tabChar17 = new TabChar();
            TabChar tabChar18 = new TabChar();
            TabChar tabChar19 = new TabChar();
            TabChar tabChar20 = new TabChar();
            TabChar tabChar21 = new TabChar();
            Text text26 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text26.Text = "/";

            run26.Append(runProperties26);
            run26.Append(tabChar12);
            run26.Append(tabChar13);
            run26.Append(tabChar14);
            run26.Append(tabChar15);
            run26.Append(tabChar16);
            run26.Append(tabChar17);
            run26.Append(tabChar18);
            run26.Append(tabChar19);
            run26.Append(tabChar20);
            run26.Append(tabChar21);
            run26.Append(text26);

            Run run27 = new Run();

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts66 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color66 = new Color(){ Val = "auto" };
            Spacing spacing66 = new Spacing(){ Val = 0 };
            Position position66 = new Position(){ Val = "0" };
            FontSize fontSize66 = new FontSize(){ Val = "24" };
            Shading shading66 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties27.Append(runFonts66);
            runProperties27.Append(color66);
            runProperties27.Append(spacing66);
            runProperties27.Append(position66);
            runProperties27.Append(fontSize66);
            runProperties27.Append(shading66);
            Text text27 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text27.Text = "студент/";

            run27.Append(runProperties27);
            run27.Append(text27);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run26);
            paragraph39.Append(run27);

            Paragraph paragraph40 = new Paragraph();

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            WidowControl widowControl40 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens40 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines40 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation40 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification40 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts runFonts67 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color67 = new Color(){ Val = "auto" };
            Spacing spacing67 = new Spacing(){ Val = 0 };
            Position position67 = new Position(){ Val = "0" };
            FontSize fontSize67 = new FontSize(){ Val = "24" };
            Shading shading67 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties40.Append(runFonts67);
            paragraphMarkRunProperties40.Append(color67);
            paragraphMarkRunProperties40.Append(spacing67);
            paragraphMarkRunProperties40.Append(position67);
            paragraphMarkRunProperties40.Append(fontSize67);
            paragraphMarkRunProperties40.Append(shading67);

            paragraphProperties40.Append(widowControl40);
            paragraphProperties40.Append(suppressAutoHyphens40);
            paragraphProperties40.Append(spacingBetweenLines40);
            paragraphProperties40.Append(indentation40);
            paragraphProperties40.Append(justification40);
            paragraphProperties40.Append(paragraphMarkRunProperties40);

            paragraph40.Append(paragraphProperties40);

            Paragraph paragraph41 = new Paragraph();

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            WidowControl widowControl41 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens41 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines41 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation41 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification41 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts runFonts68 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color68 = new Color(){ Val = "auto" };
            Spacing spacing68 = new Spacing(){ Val = 0 };
            Position position68 = new Position(){ Val = "0" };
            FontSize fontSize68 = new FontSize(){ Val = "24" };
            Shading shading68 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties41.Append(runFonts68);
            paragraphMarkRunProperties41.Append(color68);
            paragraphMarkRunProperties41.Append(spacing68);
            paragraphMarkRunProperties41.Append(position68);
            paragraphMarkRunProperties41.Append(fontSize68);
            paragraphMarkRunProperties41.Append(shading68);

            paragraphProperties41.Append(widowControl41);
            paragraphProperties41.Append(suppressAutoHyphens41);
            paragraphProperties41.Append(spacingBetweenLines41);
            paragraphProperties41.Append(indentation41);
            paragraphProperties41.Append(justification41);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            paragraph41.Append(paragraphProperties41);

            Paragraph paragraph42 = new Paragraph();

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            WidowControl widowControl42 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens42 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines42 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation42 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification42 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts69 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color69 = new Color(){ Val = "auto" };
            Spacing spacing69 = new Spacing(){ Val = 0 };
            Position position69 = new Position(){ Val = "0" };
            FontSize fontSize69 = new FontSize(){ Val = "24" };
            Shading shading69 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties42.Append(runFonts69);
            paragraphMarkRunProperties42.Append(color69);
            paragraphMarkRunProperties42.Append(spacing69);
            paragraphMarkRunProperties42.Append(position69);
            paragraphMarkRunProperties42.Append(fontSize69);
            paragraphMarkRunProperties42.Append(shading69);

            paragraphProperties42.Append(widowControl42);
            paragraphProperties42.Append(suppressAutoHyphens42);
            paragraphProperties42.Append(spacingBetweenLines42);
            paragraphProperties42.Append(indentation42);
            paragraphProperties42.Append(justification42);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            Run run28 = new Run();

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts70 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color70 = new Color(){ Val = "auto" };
            Spacing spacing70 = new Spacing(){ Val = 0 };
            Position position70 = new Position(){ Val = "0" };
            FontSize fontSize70 = new FontSize(){ Val = "24" };
            Shading shading70 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties28.Append(runFonts70);
            runProperties28.Append(color70);
            runProperties28.Append(spacing70);
            runProperties28.Append(position70);
            runProperties28.Append(fontSize70);
            runProperties28.Append(shading70);
            TabChar tabChar22 = new TabChar();
            TabChar tabChar23 = new TabChar();
            TabChar tabChar24 = new TabChar();
            TabChar tabChar25 = new TabChar();
            TabChar tabChar26 = new TabChar();
            TabChar tabChar27 = new TabChar();
            TabChar tabChar28 = new TabChar();
            TabChar tabChar29 = new TabChar();
            TabChar tabChar30 = new TabChar();
            TabChar tabChar31 = new TabChar();
            Text text28 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text28.Text = "/";

            run28.Append(runProperties28);
            run28.Append(tabChar22);
            run28.Append(tabChar23);
            run28.Append(tabChar24);
            run28.Append(tabChar25);
            run28.Append(tabChar26);
            run28.Append(tabChar27);
            run28.Append(tabChar28);
            run28.Append(tabChar29);
            run28.Append(tabChar30);
            run28.Append(tabChar31);
            run28.Append(text28);

            Run run29 = new Run();

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts71 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color71 = new Color(){ Val = "auto" };
            Spacing spacing71 = new Spacing(){ Val = 0 };
            Position position71 = new Position(){ Val = "0" };
            FontSize fontSize71 = new FontSize(){ Val = "24" };
            Shading shading71 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            runProperties29.Append(runFonts71);
            runProperties29.Append(color71);
            runProperties29.Append(spacing71);
            runProperties29.Append(position71);
            runProperties29.Append(fontSize71);
            runProperties29.Append(shading71);
            Text text29 = new Text(){ Space = SpaceProcessingModeValues.Preserve };
            text29.Text = "научен р-л/";

            run29.Append(runProperties29);
            run29.Append(text29);

            paragraph42.Append(paragraphProperties42);
            paragraph42.Append(run28);
            paragraph42.Append(run29);

            Paragraph paragraph43 = new Paragraph();

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            WidowControl widowControl43 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens43 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines43 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation43 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification43 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            RunFonts runFonts72 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color72 = new Color(){ Val = "auto" };
            Spacing spacing72 = new Spacing(){ Val = 0 };
            Position position72 = new Position(){ Val = "0" };
            FontSize fontSize72 = new FontSize(){ Val = "24" };
            Shading shading72 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties43.Append(runFonts72);
            paragraphMarkRunProperties43.Append(color72);
            paragraphMarkRunProperties43.Append(spacing72);
            paragraphMarkRunProperties43.Append(position72);
            paragraphMarkRunProperties43.Append(fontSize72);
            paragraphMarkRunProperties43.Append(shading72);

            paragraphProperties43.Append(widowControl43);
            paragraphProperties43.Append(suppressAutoHyphens43);
            paragraphProperties43.Append(spacingBetweenLines43);
            paragraphProperties43.Append(indentation43);
            paragraphProperties43.Append(justification43);
            paragraphProperties43.Append(paragraphMarkRunProperties43);

            paragraph43.Append(paragraphProperties43);

            Paragraph paragraph44 = new Paragraph();

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();
            WidowControl widowControl44 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens44 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines44 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation44 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification44 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            RunFonts runFonts73 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color73 = new Color(){ Val = "auto" };
            Spacing spacing73 = new Spacing(){ Val = 0 };
            Position position73 = new Position(){ Val = "0" };
            FontSize fontSize73 = new FontSize(){ Val = "24" };
            Shading shading73 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties44.Append(runFonts73);
            paragraphMarkRunProperties44.Append(color73);
            paragraphMarkRunProperties44.Append(spacing73);
            paragraphMarkRunProperties44.Append(position73);
            paragraphMarkRunProperties44.Append(fontSize73);
            paragraphMarkRunProperties44.Append(shading73);

            paragraphProperties44.Append(widowControl44);
            paragraphProperties44.Append(suppressAutoHyphens44);
            paragraphProperties44.Append(spacingBetweenLines44);
            paragraphProperties44.Append(indentation44);
            paragraphProperties44.Append(justification44);
            paragraphProperties44.Append(paragraphMarkRunProperties44);

            paragraph44.Append(paragraphProperties44);

            Paragraph paragraph45 = new Paragraph();

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            WidowControl widowControl45 = new WidowControl(){ Val = false };
            SuppressAutoHyphens suppressAutoHyphens45 = new SuppressAutoHyphens(){ Val = true };
            SpacingBetweenLines spacingBetweenLines45 = new SpacingBetweenLines(){ Before = "0", After = "0", Line = "240" };
            Indentation indentation45 = new Indentation(){ Start = "0", End = "0", FirstLine = "0" };
            Justification justification45 = new Justification(){ Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            RunFonts runFonts74 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color74 = new Color(){ Val = "auto" };
            Spacing spacing74 = new Spacing(){ Val = 0 };
            Position position74 = new Position(){ Val = "0" };
            FontSize fontSize74 = new FontSize(){ Val = "24" };
            Shading shading74 = new Shading(){ Val = ShadingPatternValues.Clear, Fill = "auto" };

            paragraphMarkRunProperties45.Append(runFonts74);
            paragraphMarkRunProperties45.Append(color74);
            paragraphMarkRunProperties45.Append(spacing74);
            paragraphMarkRunProperties45.Append(position74);
            paragraphMarkRunProperties45.Append(fontSize74);
            paragraphMarkRunProperties45.Append(shading74);

            paragraphProperties45.Append(widowControl45);
            paragraphProperties45.Append(suppressAutoHyphens45);
            paragraphProperties45.Append(spacingBetweenLines45);
            paragraphProperties45.Append(indentation45);
            paragraphProperties45.Append(justification45);
            paragraphProperties45.Append(paragraphMarkRunProperties45);

            paragraph45.Append(paragraphProperties45);

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(paragraph3);
            body1.Append(paragraph4);
            body1.Append(paragraph5);
            body1.Append(paragraph6);
            body1.Append(paragraph7);
            body1.Append(paragraph8);
            body1.Append(paragraph9);
            body1.Append(paragraph10);
            body1.Append(paragraph11);
            body1.Append(paragraph12);
            body1.Append(paragraph13);
            body1.Append(paragraph14);
            body1.Append(paragraph15);
            body1.Append(paragraph16);
            body1.Append(paragraph17);
            body1.Append(paragraph18);
            body1.Append(paragraph19);
            body1.Append(paragraph20);
            body1.Append(paragraph21);
            body1.Append(paragraph22);
            body1.Append(paragraph23);
            body1.Append(paragraph24);
            body1.Append(paragraph25);
            body1.Append(paragraph26);
            body1.Append(paragraph27);
            body1.Append(paragraph28);
            body1.Append(paragraph29);
            body1.Append(paragraph30);
            body1.Append(paragraph31);
            body1.Append(paragraph32);
            body1.Append(paragraph33);
            body1.Append(paragraph34);
            body1.Append(paragraph35);
            body1.Append(paragraph36);
            body1.Append(paragraph37);
            body1.Append(paragraph38);
            body1.Append(paragraph39);
            body1.Append(paragraph40);
            body1.Append(paragraph41);
            body1.Append(paragraph42);
            body1.Append(paragraph43);
            body1.Append(paragraph44);
            body1.Append(paragraph45);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering();
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleDefinitionsPart1.Styles = styles1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
        }
    }
}

