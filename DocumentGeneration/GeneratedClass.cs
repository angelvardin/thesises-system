using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using System.Collections.Generic;
using DissProject.Models;

namespace DocumentGeneration
{
    public class GeneratedClass
    {
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath, IndividualPlan individualPlan)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package, individualPlan);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document, IndividualPlan individualPlan)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1, individualPlan);

            StylesWithEffectsPart stylesWithEffectsPart1 = mainDocumentPart1.AddNewPart<StylesWithEffectsPart>("rId3");
            GenerateStylesWithEffectsPart1Content(stylesWithEffectsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId7");
            GenerateThemePart1Content(themePart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId1");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId6");
            GenerateFontTablePart1Content(fontTablePart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId5");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId4");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "140";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "801";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "6";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Title";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Приложение No";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "FMI";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "940";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1, IndividualPlan individualPlan)
        {
            DocumentFormat.OpenXml.Wordprocessing.Document document1 = new DocumentFormat.OpenXml.Wordprocessing.Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "0034307B", RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            FontSize fontSize1 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };
            Languages languages1 = new Languages() { Val = "ru-RU" };

            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties1.Append(languages1);

            paragraphProperties1.Append(paragraphMarkRunProperties1);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            FontSize fontSize2 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = "Приложение ";

            run1.Append(runProperties1);
            run1.Append(text1);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            FontSize fontSize3 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };
            Languages languages2 = new Languages() { Val = "en-US" };

            runProperties2.Append(fontSize3);
            runProperties2.Append(fontSizeComplexScript3);
            runProperties2.Append(languages2);
            Text text2 = new Text();
            text2.Text = "No";

            run2.Append(runProperties2);
            run2.Append(text2);

            Run run3 = new Run() { RsidRunProperties = "0034307B" };

            RunProperties runProperties3 = new RunProperties();
            FontSize fontSize4 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };
            Languages languages3 = new Languages() { Val = "ru-RU" };

            runProperties3.Append(fontSize4);
            runProperties3.Append(fontSizeComplexScript4);
            runProperties3.Append(languages3);
            Text text3 = new Text();
            text3.Text = ".2";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(bookmarkStart1);
            paragraph1.Append(bookmarkEnd1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "0034307B", RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            FontSize fontSize5 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };
            Languages languages4 = new Languages() { Val = "ru-RU" };

            paragraphMarkRunProperties2.Append(fontSize5);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript5);
            paragraphMarkRunProperties2.Append(languages4);

            paragraphProperties2.Append(paragraphMarkRunProperties2);

            paragraph2.Append(paragraphProperties2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            FontSize fontSize6 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties3.Append(fontSize6);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript6);

            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            FontSize fontSize7 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

            runProperties4.Append(fontSize7);
            runProperties4.Append(fontSizeComplexScript7);
            Text text4 = new Text();
            text4.Text = "СУ “С";

            run4.Append(runProperties4);
            run4.Append(text4);

            Run run5 = new Run() { RsidRunAddition = "0039086C" };

            RunProperties runProperties5 = new RunProperties();
            FontSize fontSize8 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

            runProperties5.Append(fontSize8);
            runProperties5.Append(fontSizeComplexScript8);
            Text text5 = new Text();
            text5.Text = "в";

            run5.Append(runProperties5);
            run5.Append(text5);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            FontSize fontSize9 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

            runProperties6.Append(fontSize9);
            runProperties6.Append(fontSizeComplexScript9);
            Text text6 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text6.Text = ".Климент Охридски”  ";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run4);
            paragraph3.Append(run5);
            paragraph3.Append(run6);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "00D2157D" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            FontSize fontSize10 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties4.Append(fontSize10);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript10);

            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            FontSize fontSize11 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

            runProperties7.Append(fontSize11);
            runProperties7.Append(fontSizeComplexScript11);
            Text text7 = new Text();
            text7.Text = "Факултет по математика и информатика";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run7);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            FontSize fontSize12 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties5.Append(fontSize12);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript12);

            paragraphProperties5.Append(paragraphMarkRunProperties5);

            paragraph5.Append(paragraphProperties5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            FontSize fontSize13 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties6.Append(fontSize13);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript13);

            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            FontSize fontSize14 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "28" };

            runProperties8.Append(fontSize14);
            runProperties8.Append(fontSizeComplexScript14);
            Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text8.Text = "докторантура: ";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run() { RsidRunAddition = "00874051" };

            RunProperties runProperties9 = new RunProperties();
            FontSize fontSize15 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "28" };

            runProperties9.Append(fontSize15);
            runProperties9.Append(fontSizeComplexScript15);
            Text text9 = new Text();
            text9.Text = individualPlan.PhdStudent.FormOfEducation.ToString();

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run8);
            paragraph6.Append(run9);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            FontSize fontSize16 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties7.Append(fontSize16);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript16);

            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            FontSize fontSize17 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "28" };

            runProperties10.Append(fontSize17);
            runProperties10.Append(fontSizeComplexScript17);
            Text text10 = new Text();
            text10.Text = "ФМИ";

            run10.Append(runProperties10);
            run10.Append(text10);

            Run run11 = new Run() { RsidRunAddition = "00874051" };

            RunProperties runProperties11 = new RunProperties();
            FontSize fontSize18 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "28" };

            runProperties11.Append(fontSize18);
            runProperties11.Append(fontSizeComplexScript18);
            Text text11 = new Text();
            text11.Text = ", К";

            run11.Append(runProperties11);
            run11.Append(text11);

            Run run12 = new Run() { RsidRunAddition = "00D2157D" };

            RunProperties runProperties12 = new RunProperties();
            FontSize fontSize19 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "28" };

            runProperties12.Append(fontSize19);
            runProperties12.Append(fontSizeComplexScript19);
            Text text12 = new Text();
            text12.Text = "атедра \"" +individualPlan.PhdStudent.Department.Description + "\"";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run10);
            paragraph7.Append(run11);
            paragraph7.Append(run12);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            FontSize fontSize20 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties8.Append(fontSize20);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript20);

            paragraphProperties8.Append(paragraphMarkRunProperties8);

            paragraph8.Append(paragraphProperties8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "00874051", RsidRunAdditionDefault = "00874051" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            FontSize fontSize21 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties9.Append(fontSize21);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript21);

            paragraphProperties9.Append(paragraphMarkRunProperties9);

            paragraph9.Append(paragraphProperties9);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00870F43", RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize22 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties10.Append(bold1);
            paragraphMarkRunProperties10.Append(boldComplexScript1);
            paragraphMarkRunProperties10.Append(fontSize22);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript22);

            paragraphProperties10.Append(justification1);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            paragraph10.Append(paragraphProperties10);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize23 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties11.Append(bold2);
            paragraphMarkRunProperties11.Append(boldComplexScript2);
            paragraphMarkRunProperties11.Append(fontSize23);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript23);

            paragraphProperties11.Append(justification2);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize24 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "28" };

            runProperties13.Append(bold3);
            runProperties13.Append(boldComplexScript3);
            runProperties13.Append(fontSize24);
            runProperties13.Append(fontSizeComplexScript24);
            Text text13 = new Text();
            text13.Text = "И Н Д И В И Д У А Л Е Н    П Л А Н";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run13);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            FontSize fontSize25 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties12.Append(bold4);
            paragraphMarkRunProperties12.Append(boldComplexScript4);
            paragraphMarkRunProperties12.Append(fontSize25);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript25);

            paragraphProperties12.Append(justification3);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            FontSize fontSize26 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "28" };

            runProperties14.Append(bold5);
            runProperties14.Append(boldComplexScript5);
            runProperties14.Append(fontSize26);
            runProperties14.Append(fontSizeComplexScript26);
            Text text14 = new Text();
            text14.Text = "За работа на докторанта";

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run14);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            FontSize fontSize27 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties13.Append(bold6);
            paragraphMarkRunProperties13.Append(boldComplexScript6);
            paragraphMarkRunProperties13.Append(fontSize27);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript27);

            paragraphProperties13.Append(justification4);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph13.Append(paragraphProperties13);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId1 = new NumberingId() { Val = 2 };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs1.Append(tabStop1);
            Justification justification5 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            FontSize fontSize28 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties14.Append(fontSize28);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript28);

            paragraphProperties14.Append(numberingProperties1);
            paragraphProperties14.Append(tabs1);
            paragraphProperties14.Append(justification5);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            FontSize fontSize29 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "28" };

            runProperties15.Append(fontSize29);
            runProperties15.Append(fontSizeComplexScript29);
            Text text15 = new Text();
            text15.Text = "Име, презиме и фамилия";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunAddition = "00792705" };

            RunProperties runProperties16 = new RunProperties();
            FontSize fontSize30 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "28" };

            runProperties16.Append(fontSize30);
            runProperties16.Append(fontSizeComplexScript30);
            Text text16 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text16.Text = ": ";

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run() { RsidRunAddition = "008A372A" };

            RunProperties runProperties17 = new RunProperties();
            FontSize fontSize31 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "28" };
            Languages languages5 = new Languages() { Val = "en-US" };

            runProperties17.Append(fontSize31);
            runProperties17.Append(fontSizeComplexScript31);
            runProperties17.Append(languages5);
            Text text17 = new Text();
            text17.Text = individualPlan.PhdStudent.AllNames;

            run17.Append(runProperties17);
            run17.Append(text17);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run15);
            paragraph14.Append(run16);
            paragraph14.Append(run17);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();

            NumberingProperties numberingProperties2 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference2 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId2 = new NumberingId() { Val = 2 };

            numberingProperties2.Append(numberingLevelReference2);
            numberingProperties2.Append(numberingId2);

            Tabs tabs2 = new Tabs();
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs2.Append(tabStop2);
            Justification justification6 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            FontSize fontSize32 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties15.Append(fontSize32);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript32);

            paragraphProperties15.Append(numberingProperties2);
            paragraphProperties15.Append(tabs2);
            paragraphProperties15.Append(justification6);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run18 = new Run();

            RunProperties runProperties18 = new RunProperties();
            FontSize fontSize33 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "28" };

            runProperties18.Append(fontSize33);
            runProperties18.Append(fontSizeComplexScript33);
            Text text18 = new Text();
            text18.Text = "Дата на зачисляване в докторантура";

            run18.Append(runProperties18);
            run18.Append(text18);

            Run run19 = new Run() { RsidRunProperties = "00870F43" };

            RunProperties runProperties19 = new RunProperties();
            FontSize fontSize34 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "28" };

            runProperties19.Append(fontSize34);
            runProperties19.Append(fontSizeComplexScript34);
            Text text19 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text19.Text = "  ";

            run19.Append(runProperties19);
            run19.Append(text19);

            Run run20 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "00717F03" };

            RunProperties runProperties20 = new RunProperties();
            FontSize fontSize35 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "28" };

            runProperties20.Append(fontSize35);
            runProperties20.Append(fontSizeComplexScript35);
            Text text20 = new Text();
            text20.Text = individualPlan.PhdStudent.DateOfApproval.ToString("dd/MM/yyyy");

            run20.Append(runProperties20);
            run20.Append(text20);

            Run run21 = new Run() { RsidRunAddition = "00717F03" };

            RunProperties runProperties21 = new RunProperties();
            FontSize fontSize36 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "28" };

            runProperties21.Append(fontSize36);
            runProperties21.Append(fontSizeComplexScript36);
            Text text21 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text21.Text = ", ";

            run21.Append(runProperties21);
            run21.Append(text21);

            Run run22 = new Run();

            RunProperties runProperties22 = new RunProperties();
            FontSize fontSize37 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "28" };

            runProperties22.Append(fontSize37);
            runProperties22.Append(fontSizeComplexScript37);
            Text text22 = new Text();
            text22.Text = "Заповед";

            run22.Append(runProperties22);
            run22.Append(text22);

            Run run23 = new Run() { RsidRunAddition = "00343988" };

            RunProperties runProperties23 = new RunProperties();
            FontSize fontSize38 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "28" };

            runProperties23.Append(fontSize38);
            runProperties23.Append(fontSizeComplexScript38);
            Text text23 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text23.Text = " ";

            run23.Append(runProperties23);
            run23.Append(text23);

            Run run24 = new Run() { RsidRunAddition = "00717F03" };

            RunProperties runProperties24 = new RunProperties();
            FontSize fontSize39 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "28" };

            runProperties24.Append(fontSize39);
            runProperties24.Append(fontSizeComplexScript39);
            Text text24 = new Text();
            text24.Text = "№";

            run24.Append(runProperties24);
            run24.Append(text24);

            Run run25 = new Run();

            RunProperties runProperties25 = new RunProperties();
            FontSize fontSize40 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "28" };

            runProperties25.Append(fontSize40);
            runProperties25.Append(fontSizeComplexScript40);
            Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text25.Text = " ";

            run25.Append(runProperties25);
            run25.Append(text25);

            Run run26 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "006221ED" };

            RunProperties runProperties26 = new RunProperties();
            FontSize fontSize41 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "28" };

            runProperties26.Append(fontSize41);
            runProperties26.Append(fontSizeComplexScript41);
            Text text26 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text26.Text = individualPlan.PhdStudent.Protocol;

            run26.Append(runProperties26);
            run26.Append(text26);

            Run run27 = new Run();

            RunProperties runProperties27 = new RunProperties();
            FontSize fontSize42 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "28" };

            runProperties27.Append(fontSize42);
            runProperties27.Append(fontSizeComplexScript42);
            Text text27 = new Text();
            text27.Text = "от";

            run27.Append(runProperties27);
            run27.Append(text27);

            Run run28 = new Run() { RsidRunAddition = "00BA2B18" };

            RunProperties runProperties28 = new RunProperties();
            FontSize fontSize43 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "28" };

            runProperties28.Append(fontSize43);
            runProperties28.Append(fontSizeComplexScript43);
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = " ";

            run28.Append(runProperties28);
            run28.Append(text28);

            Run run29 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "006221ED" };

            RunProperties runProperties29 = new RunProperties();
            FontSize fontSize44 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "28" };

            runProperties29.Append(fontSize44);
            runProperties29.Append(fontSizeComplexScript44);
            Text text29 = new Text();
            text29.Text = "27.05.2013";

            run29.Append(runProperties29);
            run29.Append(text29);

            Run run30 = new Run() { RsidRunAddition = "00717F03" };

            RunProperties runProperties30 = new RunProperties();
            FontSize fontSize45 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "28" };

            runProperties30.Append(fontSize45);
            runProperties30.Append(fontSizeComplexScript45);
            Text text30 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text30.Text = " г.";

            run30.Append(runProperties30);
            run30.Append(text30);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run18);
            paragraph15.Append(run19);
            paragraph15.Append(run20);
            paragraph15.Append(run21);
            paragraph15.Append(run22);
            paragraph15.Append(run23);
            paragraph15.Append(run24);
            paragraph15.Append(run25);
            paragraph15.Append(run26);
            paragraph15.Append(run27);
            paragraph15.Append(run28);
            paragraph15.Append(run29);
            paragraph15.Append(run30);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();

            NumberingProperties numberingProperties3 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference3 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId3 = new NumberingId() { Val = 2 };

            numberingProperties3.Append(numberingLevelReference3);
            numberingProperties3.Append(numberingId3);

            Tabs tabs3 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs3.Append(tabStop3);
            Justification justification7 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            FontSize fontSize46 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties16.Append(fontSize46);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript46);

            paragraphProperties16.Append(numberingProperties3);
            paragraphProperties16.Append(tabs3);
            paragraphProperties16.Append(justification7);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run31 = new Run();

            RunProperties runProperties31 = new RunProperties();
            FontSize fontSize47 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "28" };

            runProperties31.Append(fontSize47);
            runProperties31.Append(fontSizeComplexScript47);
            Text text31 = new Text();
            text31.Text = "Срок на завършване на докторантурата";

            run31.Append(runProperties31);
            run31.Append(text31);

            Run run32 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties32 = new RunProperties();
            FontSize fontSize48 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "28" };

            runProperties32.Append(fontSize48);
            runProperties32.Append(fontSizeComplexScript48);
            Text text32 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text32.Text = " ";

            run32.Append(runProperties32);
            run32.Append(text32);

            Run run33 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "006221ED" };

            RunProperties runProperties33 = new RunProperties();
            FontSize fontSize49 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "28" };

            runProperties33.Append(fontSize49);
            runProperties33.Append(fontSizeComplexScript49);
            Text text33 = new Text();
            text33.Text = "20.05.";

            run33.Append(runProperties33);
            run33.Append(text33);

            Run run34 = new Run() { RsidRunAddition = "00DA4630" };

            RunProperties runProperties34 = new RunProperties();
            FontSize fontSize50 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "28" };

            runProperties34.Append(fontSize50);
            runProperties34.Append(fontSizeComplexScript50);
            Text text34 = new Text();
            text34.Text = "2016";

            run34.Append(runProperties34);
            run34.Append(text34);

            Run run35 = new Run() { RsidRunAddition = "00717F03" };

            RunProperties runProperties35 = new RunProperties();
            FontSize fontSize51 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "28" };

            runProperties35.Append(fontSize51);
            runProperties35.Append(fontSizeComplexScript51);
            Text text35 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text35.Text = " ";

            run35.Append(runProperties35);
            run35.Append(text35);

            Run run36 = new Run() { RsidRunAddition = "006221ED" };

            RunProperties runProperties36 = new RunProperties();
            FontSize fontSize52 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "28" };

            runProperties36.Append(fontSize52);
            runProperties36.Append(fontSizeComplexScript52);
            Text text36 = new Text();
            text36.Text = "г.";

            run36.Append(runProperties36);
            run36.Append(text36);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run31);
            paragraph16.Append(run32);
            paragraph16.Append(run33);
            paragraph16.Append(run34);
            paragraph16.Append(run35);
            paragraph16.Append(run36);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            NumberingProperties numberingProperties4 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference4 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId4 = new NumberingId() { Val = 2 };

            numberingProperties4.Append(numberingLevelReference4);
            numberingProperties4.Append(numberingId4);

            Tabs tabs4 = new Tabs();
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs4.Append(tabStop4);
            Justification justification8 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize53 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties17.Append(fontSize53);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript53);

            paragraphProperties17.Append(numberingProperties4);
            paragraphProperties17.Append(tabs4);
            paragraphProperties17.Append(justification8);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run37 = new Run();

            RunProperties runProperties37 = new RunProperties();
            FontSize fontSize54 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "28" };

            runProperties37.Append(fontSize54);
            runProperties37.Append(fontSizeComplexScript54);
            Text text37 = new Text();
            text37.Text = "Научна специалност";

            run37.Append(runProperties37);
            run37.Append(text37);

            Run run38 = new Run() { RsidRunAddition = "00D2157D" };

            RunProperties runProperties38 = new RunProperties();
            FontSize fontSize55 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "28" };

            runProperties38.Append(fontSize55);
            runProperties38.Append(fontSizeComplexScript55);
            Text text38 = new Text();
            text38.Text = ":";

            run38.Append(runProperties38);
            run38.Append(text38);

            Run run39 = new Run() { RsidRunAddition = "0034307B" };

            RunProperties runProperties39 = new RunProperties();
            FontSize fontSize56 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "28" };

            runProperties39.Append(fontSize56);
            runProperties39.Append(fontSizeComplexScript56);
            Text text39 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text39.Text = " ";

            run39.Append(runProperties39);
            run39.Append(text39);

            Run run40 = new Run() { RsidRunProperties = "00792705", RsidRunAddition = "00792705" };

            RunProperties runProperties40 = new RunProperties();
            FontSize fontSize57 = new FontSize() { Val = "28" };

            runProperties40.Append(fontSize57);
            Text text40 = new Text();
            text40.Text = individualPlan.Specialty;

            run40.Append(runProperties40);
            run40.Append(text40);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run37);
            paragraph17.Append(run38);
            paragraph17.Append(run39);
            paragraph17.Append(run40);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            NumberingProperties numberingProperties5 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference5 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId5 = new NumberingId() { Val = 2 };

            numberingProperties5.Append(numberingLevelReference5);
            numberingProperties5.Append(numberingId5);

            Tabs tabs5 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs5.Append(tabStop5);
            Justification justification9 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            FontSize fontSize58 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties18.Append(fontSize58);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript57);

            paragraphProperties18.Append(numberingProperties5);
            paragraphProperties18.Append(tabs5);
            paragraphProperties18.Append(justification9);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run41 = new Run();

            RunProperties runProperties41 = new RunProperties();
            FontSize fontSize59 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "28" };

            runProperties41.Append(fontSize59);
            runProperties41.Append(fontSizeComplexScript58);
            Text text41 = new Text();
            text41.Text = "Тема на дисертционната работа";

            run41.Append(runProperties41);
            run41.Append(text41);

            Run run42 = new Run() { RsidRunAddition = "00D2157D" };

            RunProperties runProperties42 = new RunProperties();
            FontSize fontSize60 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "28" };

            runProperties42.Append(fontSize60);
            runProperties42.Append(fontSizeComplexScript59);
            Text text42 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text42.Text = ": ";

            run42.Append(runProperties42);
            run42.Append(text42);

            Run run43 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "008A372A" };

            RunProperties runProperties43 = new RunProperties();
            FontSize fontSize61 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "28" };

            runProperties43.Append(fontSize61);
            runProperties43.Append(fontSizeComplexScript60);
            Text text43 = new Text();
            text43.Text = "“";

            run43.Append(runProperties43);
            run43.Append(text43);

            Run run44 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "00792705" };

            RunProperties runProperties44 = new RunProperties();
            //Italic italic1 = new Italic();
            FontSize fontSize62 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "28" };

            //runProperties44.Append(italic1);
            runProperties44.Append(fontSize62);
            runProperties44.Append(fontSizeComplexScript61);
            Text text44 = new Text();
            text44.Text = individualPlan.PhdThesisTitle;

            run44.Append(runProperties44);
            run44.Append(text44);

            //Run run45 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "008A372A" };

            //RunProperties runProperties45 = new RunProperties();
            //Italic italic2 = new Italic();
            //FontSize fontSize63 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "28" };

            //runProperties45.Append(italic2);
            //runProperties45.Append(fontSize63);
            //runProperties45.Append(fontSizeComplexScript62);
            //Text text45 = new Text();
            //text45.Text = "”";

            //run45.Append(runProperties45);
            //run45.Append(text45);

            //Run run46 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "008A372A" };

            //RunProperties runProperties46 = new RunProperties();
            //FontSize fontSize64 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "28" };

            //runProperties46.Append(fontSize64);
            //runProperties46.Append(fontSizeComplexScript63);
            //Text text46 = new Text();
            //text46.Text = ",";

            //run46.Append(runProperties46);
            //run46.Append(text46);

            //Run run47 = new Run() { RsidRunAddition = "000405CB" };

            //RunProperties runProperties47 = new RunProperties();
            //FontSize fontSize65 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "28" };
            //Languages languages6 = new Languages() { Val = "ru-RU" };

            //runProperties47.Append(fontSize65);
            //runProperties47.Append(fontSizeComplexScript64);
            //runProperties47.Append(languages6);
            //Text text47 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text47.Text = " ";

            //run47.Append(runProperties47);
            //run47.Append(text47);

            //Run run48 = new Run();

            //RunProperties runProperties48 = new RunProperties();
            //FontSize fontSize66 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "28" };

            //runProperties48.Append(fontSize66);
            //runProperties48.Append(fontSizeComplexScript65);
            //Text text48 = new Text();
            //text48.Text = "утвърдена";

            //run48.Append(runProperties48);
            //run48.Append(text48);

            //Run run49 = new Run() { RsidRunAddition = "00792705" };

            //RunProperties runProperties49 = new RunProperties();
            //FontSize fontSize67 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "28" };

            //runProperties49.Append(fontSize67);
            //runProperties49.Append(fontSizeComplexScript66);
            //Text text49 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text49.Text = " ";

            //run49.Append(runProperties49);
            //run49.Append(text49);

            //Run run50 = new Run();

            //RunProperties runProperties50 = new RunProperties();
            //FontSize fontSize68 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "28" };

            //runProperties50.Append(fontSize68);
            //runProperties50.Append(fontSizeComplexScript67);
            //Text text50 = new Text();
            //text50.Text = "от Факултетния съвет на";

            //run50.Append(runProperties50);
            //run50.Append(text50);

            //Run run51 = new Run() { RsidRunAddition = "006221ED" };

            //RunProperties runProperties51 = new RunProperties();
            //FontSize fontSize69 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "28" };

            //runProperties51.Append(fontSize69);
            //runProperties51.Append(fontSizeComplexScript68);
            //Text text51 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text51.Text = " Факултета по математика и информатика на";

            //run51.Append(runProperties51);
            //run51.Append(text51);

            //Run run52 = new Run();

            //RunProperties runProperties52 = new RunProperties();
            //FontSize fontSize70 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "28" };

            //runProperties52.Append(fontSize70);
            //runProperties52.Append(fontSizeComplexScript69);
            //Text text52 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text52.Text = " зас";

            //run52.Append(runProperties52);
            //run52.Append(text52);

            //Run run53 = new Run() { RsidRunAddition = "006221ED" };

            //RunProperties runProperties53 = new RunProperties();
            //FontSize fontSize71 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "28" };

            //runProperties53.Append(fontSize71);
            //runProperties53.Append(fontSizeComplexScript70);
            //Text text53 = new Text();
            //text53.Text = "едание от 29.04.2013";

            //run53.Append(runProperties53);
            //run53.Append(text53);

            //Run run54 = new Run() { RsidRunAddition = "00717F03" };

            //RunProperties runProperties54 = new RunProperties();
            //FontSize fontSize72 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "28" };

            //runProperties54.Append(fontSize72);
            //runProperties54.Append(fontSizeComplexScript71);
            //Text text54 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text54.Text = " ";

            //run54.Append(runProperties54);
            //run54.Append(text54);

            //Run run55 = new Run() { RsidRunAddition = "006221ED" };

            //RunProperties runProperties55 = new RunProperties();
            //FontSize fontSize73 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "28" };

            //runProperties55.Append(fontSize73);
            //runProperties55.Append(fontSizeComplexScript72);
            //Text text55 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text55.Text = "г. ";

            //run55.Append(runProperties55);
            //run55.Append(text55);

            //Run run56 = new Run();

            //RunProperties runProperties56 = new RunProperties();
            //FontSize fontSize74 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "28" };

            //runProperties56.Append(fontSize74);
            //runProperties56.Append(fontSizeComplexScript73);
            //Text text56 = new Text();
            //text56.Text = "Протокол";

            //run56.Append(runProperties56);
            //run56.Append(text56);

            //Run run57 = new Run() { RsidRunAddition = "00717F03" };

            //RunProperties runProperties57 = new RunProperties();
            //FontSize fontSize75 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "28" };

            //runProperties57.Append(fontSize75);
            //runProperties57.Append(fontSizeComplexScript74);
            //Text text57 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text57.Text = " №";

            //run57.Append(runProperties57);
            //run57.Append(text57);

            //Run run58 = new Run() { RsidRunAddition = "006221ED" };

            //RunProperties runProperties58 = new RunProperties();
            //FontSize fontSize76 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "28" };

            //runProperties58.Append(fontSize76);
            //runProperties58.Append(fontSizeComplexScript75);
            //Text text58 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text58.Text = " 4 и доклад на декана с вх.";

            //run58.Append(runProperties58);
            //run58.Append(text58);

            //Run run59 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "006221ED" };

            //RunProperties runProperties59 = new RunProperties();
            //FontSize fontSize77 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "28" };

            //runProperties59.Append(fontSize77);
            //runProperties59.Append(fontSizeComplexScript76);
            //Text text59 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text59.Text = " ";

            //run59.Append(runProperties59);
            //run59.Append(text59);

            //Run run60 = new Run() { RsidRunAddition = "00717F03" };

            //RunProperties runProperties60 = new RunProperties();
            //FontSize fontSize78 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "28" };

            //runProperties60.Append(fontSize78);
            //runProperties60.Append(fontSizeComplexScript77);
            //Text text60 = new Text();
            //text60.Text = "№";

            //run60.Append(runProperties60);
            //run60.Append(text60);

            //Run run61 = new Run() { RsidRunAddition = "006221ED" };

            //RunProperties runProperties61 = new RunProperties();
            //FontSize fontSize79 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "28" };

            //runProperties61.Append(fontSize79);
            //runProperties61.Append(fontSizeComplexScript78);
            //Text text61 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text61.Text = " 70.05.-339/17.05.2013";

            //run61.Append(runProperties61);
            //run61.Append(text61);

            //Run run62 = new Run() { RsidRunAddition = "00717F03" };

            //RunProperties runProperties62 = new RunProperties();
            //FontSize fontSize80 = new FontSize() { Val = "28" };
            //FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "28" };

            //runProperties62.Append(fontSize80);
            //runProperties62.Append(fontSizeComplexScript79);
            //Text text62 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text62.Text = " ";

            //run62.Append(runProperties62);
            //run62.Append(text62);

            Run run63 = new Run() { RsidRunAddition = "006221ED" };

            RunProperties runProperties63 = new RunProperties();
            FontSize fontSize81 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "28" };

            runProperties63.Append(fontSize81);
            runProperties63.Append(fontSizeComplexScript80);
            Text text63 = new Text();
            text63.Text = "г.";

            run63.Append(runProperties63);
            run63.Append(text63);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run41);
            paragraph18.Append(run42);
            paragraph18.Append(run43);
            paragraph18.Append(run44);
            //paragraph18.Append(run45);
            //paragraph18.Append(run46);
            //paragraph18.Append(run47);
            //paragraph18.Append(run48);
            //paragraph18.Append(run49);
            //paragraph18.Append(run50);
            //paragraph18.Append(run51);
            //paragraph18.Append(run52);
            //paragraph18.Append(run53);
            //paragraph18.Append(run54);
            //paragraph18.Append(run55);
            //paragraph18.Append(run56);
            //paragraph18.Append(run57);
            //paragraph18.Append(run58);
            //paragraph18.Append(run59);
            //paragraph18.Append(run60);
            //paragraph18.Append(run61);
            //paragraph18.Append(run62);
            paragraph18.Append(run63);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "00D2157D", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();

            NumberingProperties numberingProperties6 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference6 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId6 = new NumberingId() { Val = 2 };

            numberingProperties6.Append(numberingLevelReference6);
            numberingProperties6.Append(numberingId6);

            Tabs tabs6 = new Tabs();
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs6.Append(tabStop6);

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            FontSize fontSize82 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties19.Append(fontSize82);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript81);

            paragraphProperties19.Append(numberingProperties6);
            paragraphProperties19.Append(tabs6);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run64 = new Run();

            RunProperties runProperties64 = new RunProperties();
            FontSize fontSize83 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "28" };

            runProperties64.Append(fontSize83);
            runProperties64.Append(fontSizeComplexScript82);
            Text text64 = new Text();
            text64.Text = "Научен";

            run64.Append(runProperties64);
            run64.Append(text64);

            Run run65 = new Run();

            RunProperties runProperties65 = new RunProperties();
            FontSize fontSize84 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "28" };
            Languages languages7 = new Languages() { Val = "ru-RU" };

            runProperties65.Append(fontSize84);
            runProperties65.Append(fontSizeComplexScript83);
            runProperties65.Append(languages7);
            Text text65 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text65.Text = " ";

            run65.Append(runProperties65);
            run65.Append(text65);

            Run run66 = new Run();

            RunProperties runProperties66 = new RunProperties();
            FontSize fontSize85 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "28" };

            runProperties66.Append(fontSize85);
            runProperties66.Append(fontSizeComplexScript84);
            Text text66 = new Text();
            text66.Text = "ръководител:";

            run66.Append(runProperties66);
            run66.Append(text66);

            Run run67 = new Run() { RsidRunAddition = "00D2157D" };

            RunProperties runProperties67 = new RunProperties();
            FontSize fontSize86 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "28" };

            runProperties67.Append(fontSize86);
            runProperties67.Append(fontSizeComplexScript85);
            Text text67 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text67.Text = " ";

            run67.Append(runProperties67);
            run67.Append(text67);

            Run run68 = new Run() { RsidRunAddition = "008A372A" };

            RunProperties runProperties68 = new RunProperties();
            FontSize fontSize87 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "28" };
            Languages languages8 = new Languages() { Val = "en-US" };

            runProperties68.Append(fontSize87);
            runProperties68.Append(fontSizeComplexScript86);
            runProperties68.Append(languages8);
            Text text68 = new Text();
            text68.Text = "….";

            run68.Append(runProperties68);
            run68.Append(text68);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run64);
            paragraph19.Append(run65);
            paragraph19.Append(run66);
            paragraph19.Append(run67);
            paragraph19.Append(run68);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidParagraphProperties = "00D2157D", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();

            Tabs tabs7 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9639 };

            tabs7.Append(tabStop7);
            Indentation indentation1 = new Indentation() { Start = "360" };
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            FontSize fontSize88 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties20.Append(fontSize88);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript87);

            paragraphProperties20.Append(tabs7);
            paragraphProperties20.Append(indentation1);
            paragraphProperties20.Append(justification10);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run69 = new Run();

            RunProperties runProperties69 = new RunProperties();
            FontSize fontSize89 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "28" };

            runProperties69.Append(fontSize89);
            runProperties69.Append(fontSizeComplexScript88);
            Text text69 = new Text();
            text69.Text = "(научно звание и степен, име, бащино и фамилно име)";

            run69.Append(runProperties69);
            run69.Append(text69);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run69);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "0093085C", RsidRunAdditionDefault = "0093085C" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();

            NumberingProperties numberingProperties7 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference7 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId7 = new NumberingId() { Val = 2 };

            numberingProperties7.Append(numberingLevelReference7);
            numberingProperties7.Append(numberingId7);
            Justification justification11 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            FontSize fontSize90 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties21.Append(fontSize90);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript89);

            paragraphProperties21.Append(numberingProperties7);
            paragraphProperties21.Append(justification11);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run70 = new Run();

            RunProperties runProperties70 = new RunProperties();
            FontSize fontSize91 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "28" };

            runProperties70.Append(fontSize91);
            runProperties70.Append(fontSizeComplexScript90);
            Text text70 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text70.Text = "Индивидуален план за ";

            run70.Append(runProperties70);
            run70.Append(text70);

            Run run71 = new Run() { RsidRunAddition = "00870F43" };

            RunProperties runProperties71 = new RunProperties();
            FontSize fontSize92 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "28" };

            runProperties71.Append(fontSize92);
            runProperties71.Append(fontSizeComplexScript91);
            Text text71 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text71.Text = "работа на докторанта от ";

            run71.Append(runProperties71);
            run71.Append(text71);

            Run run72 = new Run() { RsidRunAddition = "00870F43" };

            RunProperties runProperties72 = new RunProperties();
            FontSize fontSize93 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "28" };
            Languages languages9 = new Languages() { Val = "de-DE" };

            runProperties72.Append(fontSize93);
            runProperties72.Append(fontSizeComplexScript92);
            runProperties72.Append(languages9);
            Text text72 = new Text();
            text72.Text = "Ka";

            run72.Append(runProperties72);
            run72.Append(text72);

            Run run73 = new Run() { RsidRunAddition = "00870F43" };

            RunProperties runProperties73 = new RunProperties();
            FontSize fontSize94 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "28" };

            runProperties73.Append(fontSize94);
            runProperties73.Append(fontSizeComplexScript93);
            Text text73 = new Text();
            text73.Text = "тедре";

            run73.Append(runProperties73);
            run73.Append(text73);

            Run run74 = new Run();

            RunProperties runProperties74 = new RunProperties();
            FontSize fontSize95 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "28" };

            runProperties74.Append(fontSize95);
            runProperties74.Append(fontSizeComplexScript94);
            Text text74 = new Text();
            text74.Text = "ния съвет в заседание от";

            run74.Append(runProperties74);
            run74.Append(text74);

            Run run75 = new Run() { RsidRunAddition = "00D2157D" };

            RunProperties runProperties75 = new RunProperties();
            FontSize fontSize96 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "28" };

            runProperties75.Append(fontSize96);
            runProperties75.Append(fontSizeComplexScript95);
            Text text75 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text75.Text = " ";

            run75.Append(runProperties75);
            run75.Append(text75);

            Run run76 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "0039086C" };

            RunProperties runProperties76 = new RunProperties();
            FontSize fontSize97 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "28" };

            runProperties76.Append(fontSize97);
            runProperties76.Append(fontSizeComplexScript96);
            Text text76 = new Text();
            text76.Text = "";

            run76.Append(runProperties76);
            run76.Append(text76);

            Run run77 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties77 = new RunProperties();
            FontSize fontSize98 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "28" };

            runProperties77.Append(fontSize98);
            runProperties77.Append(fontSizeComplexScript97);
            Text text77 = new Text();
            text77.Text = "";

            run77.Append(runProperties77);
            run77.Append(text77);

            Run run78 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "0039086C" };

            RunProperties runProperties78 = new RunProperties();
            FontSize fontSize99 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "28" };

            runProperties78.Append(fontSize99);
            runProperties78.Append(fontSizeComplexScript98);
            Text text78 = new Text();
            text78.Text = individualPlan.FacultyProtocol;

            run78.Append(runProperties78);
            run78.Append(text78);

            Run run79 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties79 = new RunProperties();
            FontSize fontSize100 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "28" };

            runProperties79.Append(fontSize100);
            runProperties79.Append(fontSizeComplexScript99);
            Text text79 = new Text();
            text79.Text = "";

            run79.Append(runProperties79);
            run79.Append(text79);

            Run run80 = new Run();

            RunProperties runProperties80 = new RunProperties();
            FontSize fontSize101 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript100 = new FontSizeComplexScript() { Val = "28" };

            runProperties80.Append(fontSize101);
            runProperties80.Append(fontSizeComplexScript100);
            Text text80 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text80.Text = "";

            run80.Append(runProperties80);
            run80.Append(text80);

            Run run81 = new Run();

            RunProperties runProperties81 = new RunProperties();
            FontSize fontSize102 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript101 = new FontSizeComplexScript() { Val = "28" };
            Languages languages10 = new Languages() { Val = "en-US" };

            runProperties81.Append(fontSize102);
            runProperties81.Append(fontSizeComplexScript101);
            runProperties81.Append(languages10);
            Text text81 = new Text();
            text81.Text = "";

            run81.Append(runProperties81);
            run81.Append(text81);

            Run run82 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties82 = new RunProperties();
            FontSize fontSize103 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "28" };
            Languages languages11 = new Languages() { Val = "ru-RU" };

            runProperties82.Append(fontSize103);
            runProperties82.Append(fontSizeComplexScript102);
            runProperties82.Append(languages11);
            Text text82 = new Text();
            text82.Text = "";

            run82.Append(runProperties82);
            run82.Append(text82);

            Run run83 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "0039086C" };

            RunProperties runProperties83 = new RunProperties();
            FontSize fontSize104 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "28" };

            runProperties83.Append(fontSize104);
            runProperties83.Append(fontSizeComplexScript103);
            Text text83 = new Text();
            text83.Text = "";

            run83.Append(runProperties83);
            run83.Append(text83);

            Run run84 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties84 = new RunProperties();
            FontSize fontSize105 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "28" };
            Languages languages12 = new Languages() { Val = "ru-RU" };

            runProperties84.Append(fontSize105);
            runProperties84.Append(fontSizeComplexScript104);
            runProperties84.Append(languages12);
            Text text84 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text84.Text = "";

            run84.Append(runProperties84);
            run84.Append(text84);

            Run run85 = new Run();

            RunProperties runProperties85 = new RunProperties();
            FontSize fontSize106 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "28" };

            runProperties85.Append(fontSize106);
            runProperties85.Append(fontSizeComplexScript105);
            Text text85 = new Text();
            text85.Text = "";

            run85.Append(runProperties85);
            run85.Append(text85);

            Run run86 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties86 = new RunProperties();
            FontSize fontSize107 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript106 = new FontSizeComplexScript() { Val = "28" };

            runProperties86.Append(fontSize107);
            runProperties86.Append(fontSizeComplexScript106);
            Text text86 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text86.Text = "";

            run86.Append(runProperties86);
            run86.Append(text86);

            Run run87 = new Run() { RsidRunProperties = "008A372A", RsidRunAddition = "0039086C" };

            RunProperties runProperties87 = new RunProperties();
            FontSize fontSize108 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript107 = new FontSizeComplexScript() { Val = "28" };

            runProperties87.Append(fontSize108);
            runProperties87.Append(fontSizeComplexScript107);
            Text text87 = new Text();
            text87.Text = "";

            run87.Append(runProperties87);
            run87.Append(text87);

            Run run88 = new Run() { RsidRunAddition = "00F10ECE" };

            RunProperties runProperties88 = new RunProperties();
            FontSize fontSize109 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "28" };

            runProperties88.Append(fontSize109);
            runProperties88.Append(fontSizeComplexScript108);
            Text text88 = new Text();
            text88.Text = "";

            run88.Append(runProperties88);
            run88.Append(text88);

            Run run89 = new Run();

            RunProperties runProperties89 = new RunProperties();
            FontSize fontSize110 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "28" };

            runProperties89.Append(fontSize110);
            runProperties89.Append(fontSizeComplexScript109);
            Text text89 = new Text();
            text89.Text = "";

            run89.Append(runProperties89);
            run89.Append(text89);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run70);
            paragraph21.Append(run71);
            paragraph21.Append(run72);
            paragraph21.Append(run73);
            paragraph21.Append(run74);
            paragraph21.Append(run75);
            paragraph21.Append(run76);
            paragraph21.Append(run77);
            paragraph21.Append(run78);
            paragraph21.Append(run79);
            paragraph21.Append(run80);
            paragraph21.Append(run81);
            paragraph21.Append(run82);
            paragraph21.Append(run83);
            paragraph21.Append(run84);
            paragraph21.Append(run85);
            paragraph21.Append(run86);
            paragraph21.Append(run87);
            paragraph21.Append(run88);
            paragraph21.Append(run89);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00870F43", RsidParagraphAddition = "00870F43", RsidRunAdditionDefault = "00870F43" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            FontSize fontSize111 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties22.Append(fontSize111);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript110);

            paragraphProperties22.Append(paragraphMarkRunProperties22);

            paragraph22.Append(paragraphProperties22);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "00870F43", RsidR = "00870F43" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)16838U, Height = (UInt32Value)11906U, Orient = PageOrientationValues.Landscape, Code = (UInt16Value)9U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1418, Right = (UInt32Value)1418U, Bottom = 1418, Left = (UInt32Value)1418U, Header = (UInt32Value)709U, Footer = (UInt32Value)709U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "709" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

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
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of stylesWithEffectsPart1.
        private void GenerateStylesWithEffectsPart1Content(StylesWithEffectsPart stylesWithEffectsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            styles1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            styles1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            styles1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            styles1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            styles1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            styles1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            styles1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            styles1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Languages languages13 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(languages13);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 0, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "caption", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "Title", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "Subtitle", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "Strong", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "Emphasis", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Revision", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            AutoSpaceDE autoSpaceDE1 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN1 = new AutoSpaceDN() { Val = false };

            styleParagraphProperties1.Append(autoSpaceDE1);
            styleParagraphProperties1.Append(autoSpaceDN1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize112 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "24" };
            Languages languages14 = new Languages() { Val = "bg-BG" };

            styleRunProperties1.Append(fontSize112);
            styleRunProperties1.Append(fontSizeComplexScript111);
            styleRunProperties1.Append(languages14);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            SemiHidden semiHidden1 = new SemiHidden();

            style2.Append(styleName2);
            style2.Append(semiHidden1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            SemiHidden semiHidden2 = new SemiHidden();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(semiHidden2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            SemiHidden semiHidden3 = new SemiHidden();

            style4.Append(styleName4);
            style4.Append(semiHidden3);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "BalloonText" };
            StyleName styleName5 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            SemiHidden semiHidden4 = new SemiHidden();
            Rsid rsid1 = new Rsid() { Val = "00BA2B18" };

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize113 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties2.Append(runFonts2);
            styleRunProperties2.Append(fontSize113);
            styleRunProperties2.Append(fontSizeComplexScript112);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(semiHidden4);
            style5.Append(rsid1);
            style5.Append(styleRunProperties2);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);

            stylesWithEffectsPart1.Styles = styles1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles2 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            styles2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles2.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            DocDefaults docDefaults2 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault2 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Languages languages15 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle2.Append(runFonts3);
            runPropertiesBaseStyle2.Append(languages15);

            runPropertiesDefault2.Append(runPropertiesBaseStyle2);
            ParagraphPropertiesDefault paragraphPropertiesDefault2 = new ParagraphPropertiesDefault();

            docDefaults2.Append(runPropertiesDefault2);
            docDefaults2.Append(paragraphPropertiesDefault2);

            LatentStyles latentStyles2 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 0, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Normal", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "heading 1", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "heading 2", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "heading 3", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "heading 4", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "heading 5", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "heading 6", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "heading 7", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "heading 8", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "heading 9", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "caption", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Title", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Subtitle", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Strong", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Emphasis", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Revision", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };

            latentStyles2.Append(latentStyleExceptionInfo127);
            latentStyles2.Append(latentStyleExceptionInfo128);
            latentStyles2.Append(latentStyleExceptionInfo129);
            latentStyles2.Append(latentStyleExceptionInfo130);
            latentStyles2.Append(latentStyleExceptionInfo131);
            latentStyles2.Append(latentStyleExceptionInfo132);
            latentStyles2.Append(latentStyleExceptionInfo133);
            latentStyles2.Append(latentStyleExceptionInfo134);
            latentStyles2.Append(latentStyleExceptionInfo135);
            latentStyles2.Append(latentStyleExceptionInfo136);
            latentStyles2.Append(latentStyleExceptionInfo137);
            latentStyles2.Append(latentStyleExceptionInfo138);
            latentStyles2.Append(latentStyleExceptionInfo139);
            latentStyles2.Append(latentStyleExceptionInfo140);
            latentStyles2.Append(latentStyleExceptionInfo141);
            latentStyles2.Append(latentStyleExceptionInfo142);
            latentStyles2.Append(latentStyleExceptionInfo143);
            latentStyles2.Append(latentStyleExceptionInfo144);
            latentStyles2.Append(latentStyleExceptionInfo145);
            latentStyles2.Append(latentStyleExceptionInfo146);
            latentStyles2.Append(latentStyleExceptionInfo147);
            latentStyles2.Append(latentStyleExceptionInfo148);
            latentStyles2.Append(latentStyleExceptionInfo149);
            latentStyles2.Append(latentStyleExceptionInfo150);
            latentStyles2.Append(latentStyleExceptionInfo151);
            latentStyles2.Append(latentStyleExceptionInfo152);
            latentStyles2.Append(latentStyleExceptionInfo153);
            latentStyles2.Append(latentStyleExceptionInfo154);
            latentStyles2.Append(latentStyleExceptionInfo155);
            latentStyles2.Append(latentStyleExceptionInfo156);
            latentStyles2.Append(latentStyleExceptionInfo157);
            latentStyles2.Append(latentStyleExceptionInfo158);
            latentStyles2.Append(latentStyleExceptionInfo159);
            latentStyles2.Append(latentStyleExceptionInfo160);
            latentStyles2.Append(latentStyleExceptionInfo161);
            latentStyles2.Append(latentStyleExceptionInfo162);
            latentStyles2.Append(latentStyleExceptionInfo163);
            latentStyles2.Append(latentStyleExceptionInfo164);
            latentStyles2.Append(latentStyleExceptionInfo165);
            latentStyles2.Append(latentStyleExceptionInfo166);
            latentStyles2.Append(latentStyleExceptionInfo167);
            latentStyles2.Append(latentStyleExceptionInfo168);
            latentStyles2.Append(latentStyleExceptionInfo169);
            latentStyles2.Append(latentStyleExceptionInfo170);
            latentStyles2.Append(latentStyleExceptionInfo171);
            latentStyles2.Append(latentStyleExceptionInfo172);
            latentStyles2.Append(latentStyleExceptionInfo173);
            latentStyles2.Append(latentStyleExceptionInfo174);
            latentStyles2.Append(latentStyleExceptionInfo175);
            latentStyles2.Append(latentStyleExceptionInfo176);
            latentStyles2.Append(latentStyleExceptionInfo177);
            latentStyles2.Append(latentStyleExceptionInfo178);
            latentStyles2.Append(latentStyleExceptionInfo179);
            latentStyles2.Append(latentStyleExceptionInfo180);
            latentStyles2.Append(latentStyleExceptionInfo181);
            latentStyles2.Append(latentStyleExceptionInfo182);
            latentStyles2.Append(latentStyleExceptionInfo183);
            latentStyles2.Append(latentStyleExceptionInfo184);
            latentStyles2.Append(latentStyleExceptionInfo185);
            latentStyles2.Append(latentStyleExceptionInfo186);
            latentStyles2.Append(latentStyleExceptionInfo187);
            latentStyles2.Append(latentStyleExceptionInfo188);
            latentStyles2.Append(latentStyleExceptionInfo189);
            latentStyles2.Append(latentStyleExceptionInfo190);
            latentStyles2.Append(latentStyleExceptionInfo191);
            latentStyles2.Append(latentStyleExceptionInfo192);
            latentStyles2.Append(latentStyleExceptionInfo193);
            latentStyles2.Append(latentStyleExceptionInfo194);
            latentStyles2.Append(latentStyleExceptionInfo195);
            latentStyles2.Append(latentStyleExceptionInfo196);
            latentStyles2.Append(latentStyleExceptionInfo197);
            latentStyles2.Append(latentStyleExceptionInfo198);
            latentStyles2.Append(latentStyleExceptionInfo199);
            latentStyles2.Append(latentStyleExceptionInfo200);
            latentStyles2.Append(latentStyleExceptionInfo201);
            latentStyles2.Append(latentStyleExceptionInfo202);
            latentStyles2.Append(latentStyleExceptionInfo203);
            latentStyles2.Append(latentStyleExceptionInfo204);
            latentStyles2.Append(latentStyleExceptionInfo205);
            latentStyles2.Append(latentStyleExceptionInfo206);
            latentStyles2.Append(latentStyleExceptionInfo207);
            latentStyles2.Append(latentStyleExceptionInfo208);
            latentStyles2.Append(latentStyleExceptionInfo209);
            latentStyles2.Append(latentStyleExceptionInfo210);
            latentStyles2.Append(latentStyleExceptionInfo211);
            latentStyles2.Append(latentStyleExceptionInfo212);
            latentStyles2.Append(latentStyleExceptionInfo213);
            latentStyles2.Append(latentStyleExceptionInfo214);
            latentStyles2.Append(latentStyleExceptionInfo215);
            latentStyles2.Append(latentStyleExceptionInfo216);
            latentStyles2.Append(latentStyleExceptionInfo217);
            latentStyles2.Append(latentStyleExceptionInfo218);
            latentStyles2.Append(latentStyleExceptionInfo219);
            latentStyles2.Append(latentStyleExceptionInfo220);
            latentStyles2.Append(latentStyleExceptionInfo221);
            latentStyles2.Append(latentStyleExceptionInfo222);
            latentStyles2.Append(latentStyleExceptionInfo223);
            latentStyles2.Append(latentStyleExceptionInfo224);
            latentStyles2.Append(latentStyleExceptionInfo225);
            latentStyles2.Append(latentStyleExceptionInfo226);
            latentStyles2.Append(latentStyleExceptionInfo227);
            latentStyles2.Append(latentStyleExceptionInfo228);
            latentStyles2.Append(latentStyleExceptionInfo229);
            latentStyles2.Append(latentStyleExceptionInfo230);
            latentStyles2.Append(latentStyleExceptionInfo231);
            latentStyles2.Append(latentStyleExceptionInfo232);
            latentStyles2.Append(latentStyleExceptionInfo233);
            latentStyles2.Append(latentStyleExceptionInfo234);
            latentStyles2.Append(latentStyleExceptionInfo235);
            latentStyles2.Append(latentStyleExceptionInfo236);
            latentStyles2.Append(latentStyleExceptionInfo237);
            latentStyles2.Append(latentStyleExceptionInfo238);
            latentStyles2.Append(latentStyleExceptionInfo239);
            latentStyles2.Append(latentStyleExceptionInfo240);
            latentStyles2.Append(latentStyleExceptionInfo241);
            latentStyles2.Append(latentStyleExceptionInfo242);
            latentStyles2.Append(latentStyleExceptionInfo243);
            latentStyles2.Append(latentStyleExceptionInfo244);
            latentStyles2.Append(latentStyleExceptionInfo245);
            latentStyles2.Append(latentStyleExceptionInfo246);
            latentStyles2.Append(latentStyleExceptionInfo247);
            latentStyles2.Append(latentStyleExceptionInfo248);
            latentStyles2.Append(latentStyleExceptionInfo249);
            latentStyles2.Append(latentStyleExceptionInfo250);
            latentStyles2.Append(latentStyleExceptionInfo251);
            latentStyles2.Append(latentStyleExceptionInfo252);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName6 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() { Val = false };

            styleParagraphProperties2.Append(autoSpaceDE2);
            styleParagraphProperties2.Append(autoSpaceDN2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            FontSize fontSize114 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "24" };
            Languages languages16 = new Languages() { Val = "bg-BG" };

            styleRunProperties3.Append(fontSize114);
            styleRunProperties3.Append(fontSizeComplexScript113);
            styleRunProperties3.Append(languages16);

            style6.Append(styleName6);
            style6.Append(primaryStyle2);
            style6.Append(styleParagraphProperties2);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName7 = new StyleName() { Val = "Default Paragraph Font" };
            SemiHidden semiHidden5 = new SemiHidden();

            style7.Append(styleName7);
            style7.Append(semiHidden5);

            Style style8 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName8 = new StyleName() { Val = "Normal Table" };
            SemiHidden semiHidden6 = new SemiHidden();

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin2);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableCellMarginDefault2);

            style8.Append(styleName8);
            style8.Append(semiHidden6);
            style8.Append(styleTableProperties2);

            Style style9 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName9 = new StyleName() { Val = "No List" };
            SemiHidden semiHidden7 = new SemiHidden();

            style9.Append(styleName9);
            style9.Append(semiHidden7);

            Style style10 = new Style() { Type = StyleValues.Paragraph, StyleId = "BalloonText" };
            StyleName styleName10 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            SemiHidden semiHidden8 = new SemiHidden();
            Rsid rsid2 = new Rsid() { Val = "00BA2B18" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize115 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties4.Append(runFonts4);
            styleRunProperties4.Append(fontSize115);
            styleRunProperties4.Append(fontSizeComplexScript114);

            style10.Append(styleName10);
            style10.Append(basedOn2);
            style10.Append(semiHidden8);
            style10.Append(rsid2);
            style10.Append(styleRunProperties4);

            styles2.Append(docDefaults2);
            styles2.Append(latentStyles2);
            styles2.Append(style6);
            styles2.Append(style7);
            styles2.Append(style8);
            styles2.Append(style9);
            styles2.Append(style10);

            styleDefinitionsPart1.Styles = styles2;
        }

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            numbering1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            numbering1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            numbering1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            numbering1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            numbering1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            Nsid nsid1 = new Nsid() { Val = "06AD43C0" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "A61E3D14" };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

            Tabs tabs8 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

            tabs8.Append(tabStop8);
            Indentation indentation2 = new Indentation() { Start = "720", Hanging = "360" };

            previousParagraphProperties1.Append(tabs8);
            previousParagraphProperties1.Append(indentation2);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts5 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties1.Append(runFonts5);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText2 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

            Tabs tabs9 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

            tabs9.Append(tabStop9);
            Indentation indentation3 = new Indentation() { Start = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(tabs9);
            previousParagraphProperties2.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts6 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties2.Append(runFonts6);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

            Tabs tabs10 = new Tabs();
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

            tabs10.Append(tabStop10);
            Indentation indentation4 = new Indentation() { Start = "2160", Hanging = "180" };

            previousParagraphProperties3.Append(tabs10);
            previousParagraphProperties3.Append(indentation4);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts7 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties3.Append(runFonts7);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);

            Level level4 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

            Tabs tabs11 = new Tabs();
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs11.Append(tabStop11);
            Indentation indentation5 = new Indentation() { Start = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(tabs11);
            previousParagraphProperties4.Append(indentation5);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts8 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties4.Append(runFonts8);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Level level5 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText5 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

            Tabs tabs12 = new Tabs();
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

            tabs12.Append(tabStop12);
            Indentation indentation6 = new Indentation() { Start = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(tabs12);
            previousParagraphProperties5.Append(indentation6);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts9 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties5.Append(runFonts9);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);
            level5.Append(numberingSymbolRunProperties5);

            Level level6 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

            Tabs tabs13 = new Tabs();
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Number, Position = 4320 };

            tabs13.Append(tabStop13);
            Indentation indentation7 = new Indentation() { Start = "4320", Hanging = "180" };

            previousParagraphProperties6.Append(tabs13);
            previousParagraphProperties6.Append(indentation7);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts10 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties6.Append(runFonts10);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);
            level6.Append(numberingSymbolRunProperties6);

            Level level7 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

            Tabs tabs14 = new Tabs();
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Number, Position = 5040 };

            tabs14.Append(tabStop14);
            Indentation indentation8 = new Indentation() { Start = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(tabs14);
            previousParagraphProperties7.Append(indentation8);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts11 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties7.Append(runFonts11);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);
            level7.Append(numberingSymbolRunProperties7);

            Level level8 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText8 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

            Tabs tabs15 = new Tabs();
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Number, Position = 5760 };

            tabs15.Append(tabStop15);
            Indentation indentation9 = new Indentation() { Start = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(tabs15);
            previousParagraphProperties8.Append(indentation9);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts12 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties8.Append(runFonts12);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);
            level8.Append(numberingSymbolRunProperties8);

            Level level9 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

            Tabs tabs16 = new Tabs();
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Number, Position = 6480 };

            tabs16.Append(tabStop16);
            Indentation indentation10 = new Indentation() { Start = "6480", Hanging = "180" };

            previousParagraphProperties9.Append(tabs16);
            previousParagraphProperties9.Append(indentation10);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts13 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties9.Append(runFonts13);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);
            level9.Append(numberingSymbolRunProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            AbstractNum abstractNum2 = new AbstractNum() { AbstractNumberId = 1 };
            Nsid nsid2 = new Nsid() { Val = "788B35F5" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "7C2883EA" };

            Level level10 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText10 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();

            Tabs tabs17 = new Tabs();
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

            tabs17.Append(tabStop17);
            Indentation indentation11 = new Indentation() { Start = "720", Hanging = "360" };

            previousParagraphProperties10.Append(tabs17);
            previousParagraphProperties10.Append(indentation11);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts14 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties10.Append(runFonts14);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties10);

            Level level11 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText11 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();

            Tabs tabs18 = new Tabs();
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

            tabs18.Append(tabStop18);
            Indentation indentation12 = new Indentation() { Start = "1440", Hanging = "360" };

            previousParagraphProperties11.Append(tabs18);
            previousParagraphProperties11.Append(indentation12);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts15 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties11.Append(runFonts15);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);
            level11.Append(numberingSymbolRunProperties11);

            Level level12 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText12 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();

            Tabs tabs19 = new Tabs();
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

            tabs19.Append(tabStop19);
            Indentation indentation13 = new Indentation() { Start = "2160", Hanging = "180" };

            previousParagraphProperties12.Append(tabs19);
            previousParagraphProperties12.Append(indentation13);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts16 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties12.Append(runFonts16);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);
            level12.Append(numberingSymbolRunProperties12);

            Level level13 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText13 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification13 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();

            Tabs tabs20 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs20.Append(tabStop20);
            Indentation indentation14 = new Indentation() { Start = "2880", Hanging = "360" };

            previousParagraphProperties13.Append(tabs20);
            previousParagraphProperties13.Append(indentation14);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts17 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties13.Append(runFonts17);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);
            level13.Append(numberingSymbolRunProperties13);

            Level level14 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText14 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification14 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();

            Tabs tabs21 = new Tabs();
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

            tabs21.Append(tabStop21);
            Indentation indentation15 = new Indentation() { Start = "3600", Hanging = "360" };

            previousParagraphProperties14.Append(tabs21);
            previousParagraphProperties14.Append(indentation15);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts18 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties14.Append(runFonts18);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);
            level14.Append(numberingSymbolRunProperties14);

            Level level15 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText15 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification15 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();

            Tabs tabs22 = new Tabs();
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Number, Position = 4320 };

            tabs22.Append(tabStop22);
            Indentation indentation16 = new Indentation() { Start = "4320", Hanging = "180" };

            previousParagraphProperties15.Append(tabs22);
            previousParagraphProperties15.Append(indentation16);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts19 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties15.Append(runFonts19);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);
            level15.Append(numberingSymbolRunProperties15);

            Level level16 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText16 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification16 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();

            Tabs tabs23 = new Tabs();
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Number, Position = 5040 };

            tabs23.Append(tabStop23);
            Indentation indentation17 = new Indentation() { Start = "5040", Hanging = "360" };

            previousParagraphProperties16.Append(tabs23);
            previousParagraphProperties16.Append(indentation17);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts20 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties16.Append(runFonts20);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);
            level16.Append(numberingSymbolRunProperties16);

            Level level17 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText17 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification17 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();

            Tabs tabs24 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Number, Position = 5760 };

            tabs24.Append(tabStop24);
            Indentation indentation18 = new Indentation() { Start = "5760", Hanging = "360" };

            previousParagraphProperties17.Append(tabs24);
            previousParagraphProperties17.Append(indentation18);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts21 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties17.Append(runFonts21);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);
            level17.Append(numberingSymbolRunProperties17);

            Level level18 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText18 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification18 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();

            Tabs tabs25 = new Tabs();
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Number, Position = 6480 };

            tabs25.Append(tabStop25);
            Indentation indentation19 = new Indentation() { Start = "6480", Hanging = "180" };

            previousParagraphProperties18.Append(tabs25);
            previousParagraphProperties18.Append(indentation19);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts22 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties18.Append(runFonts22);

            level18.Append(startNumberingValue18);
            level18.Append(numberingFormat18);
            level18.Append(levelText18);
            level18.Append(levelJustification18);
            level18.Append(previousParagraphProperties18);
            level18.Append(numberingSymbolRunProperties18);

            abstractNum2.Append(nsid2);
            abstractNum2.Append(multiLevelType2);
            abstractNum2.Append(templateCode2);
            abstractNum2.Append(level10);
            abstractNum2.Append(level11);
            abstractNum2.Append(level12);
            abstractNum2.Append(level13);
            abstractNum2.Append(level14);
            abstractNum2.Append(level15);
            abstractNum2.Append(level16);
            abstractNum2.Append(level17);
            abstractNum2.Append(level18);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 1 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 0 };

            numberingInstance2.Append(abstractNumId2);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Font font1 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Tahoma" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "020B0604030504040204" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E1002EFF", UnicodeSignature1 = "C000605B", UnicodeSignature2 = "00000029", UnicodeSignature3 = "00000000", CodePageSignature0 = "000101FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            TargetScreenSize targetScreenSize1 = new TargetScreenSize() { Val = TargetScreenSizeValues.Sz800x600 };

            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(targetScreenSize1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "80" };
            EmbedSystemFonts embedSystemFonts1 = new EmbedSystemFonts();
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            StylePaneFormatFilter stylePaneFormatFilter1 = new StylePaneFormatFilter() { Val = "3F01", AllStyles = true, CustomStyles = false, LatentStyles = false, StylesInUse = false, HeadingStyles = false, NumberingStyles = false, TableStyles = false, DirectFormattingOnRuns = true, DirectFormattingOnParagraphs = true, DirectFormattingOnNumbering = true, DirectFormattingOnTables = true, ClearFormatting = true, Top3HeadingStyles = true, VisibleStyles = false, AlternateStyleNames = false };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 709 };
            HyphenationZone hyphenationZone1 = new HyphenationZone() { Val = "425" };
            DoNotHyphenateCaps doNotHyphenateCaps1 = new DoNotHyphenateCaps();
            DrawingGridHorizontalSpacing drawingGridHorizontalSpacing1 = new DrawingGridHorizontalSpacing() { Val = "120" };
            DrawingGridVerticalSpacing drawingGridVerticalSpacing1 = new DrawingGridVerticalSpacing() { Val = "120" };
            DisplayHorizontalDrawingGrid displayHorizontalDrawingGrid1 = new DisplayHorizontalDrawingGrid() { Val = 0 };
            DisplayVerticalDrawingGrid displayVerticalDrawingGrid1 = new DisplayVerticalDrawingGrid() { Val = 3 };
            DoNotUseMarginsForDrawingGridOrigin doNotUseMarginsForDrawingGridOrigin1 = new DoNotUseMarginsForDrawingGridOrigin();
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.CompressPunctuation };
            DoNotValidateAgainstSchema doNotValidateAgainstSchema1 = new DoNotValidateAgainstSchema();
            DoNotDemarcateInvalidXml doNotDemarcateInvalidXml1 = new DoNotDemarcateInvalidXml();

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "14" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "0093085C" };
            Rsid rsid3 = new Rsid() { Val = "00024CB9" };
            Rsid rsid4 = new Rsid() { Val = "000405CB" };
            Rsid rsid5 = new Rsid() { Val = "0025121C" };
            Rsid rsid6 = new Rsid() { Val = "003019E6" };
            Rsid rsid7 = new Rsid() { Val = "0034307B" };
            Rsid rsid8 = new Rsid() { Val = "00343988" };
            Rsid rsid9 = new Rsid() { Val = "0039086C" };
            Rsid rsid10 = new Rsid() { Val = "00406567" };
            Rsid rsid11 = new Rsid() { Val = "005239DA" };
            Rsid rsid12 = new Rsid() { Val = "006221ED" };
            Rsid rsid13 = new Rsid() { Val = "006225E9" };
            Rsid rsid14 = new Rsid() { Val = "00717F03" };
            Rsid rsid15 = new Rsid() { Val = "00792705" };
            Rsid rsid16 = new Rsid() { Val = "008347EB" };
            Rsid rsid17 = new Rsid() { Val = "00847AD9" };
            Rsid rsid18 = new Rsid() { Val = "00870F43" };
            Rsid rsid19 = new Rsid() { Val = "00874051" };
            Rsid rsid20 = new Rsid() { Val = "008A372A" };
            Rsid rsid21 = new Rsid() { Val = "0093085C" };
            Rsid rsid22 = new Rsid() { Val = "0097028B" };
            Rsid rsid23 = new Rsid() { Val = "009733F6" };
            Rsid rsid24 = new Rsid() { Val = "00977D94" };
            Rsid rsid25 = new Rsid() { Val = "00A27815" };
            Rsid rsid26 = new Rsid() { Val = "00AF25B9" };
            Rsid rsid27 = new Rsid() { Val = "00BA2B18" };
            Rsid rsid28 = new Rsid() { Val = "00D2157D" };
            Rsid rsid29 = new Rsid() { Val = "00DA4630" };
            Rsid rsid30 = new Rsid() { Val = "00F10ECE" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);
            rsids1.Append(rsid13);
            rsids1.Append(rsid14);
            rsids1.Append(rsid15);
            rsids1.Append(rsid16);
            rsids1.Append(rsid17);
            rsids1.Append(rsid18);
            rsids1.Append(rsid19);
            rsids1.Append(rsid20);
            rsids1.Append(rsid21);
            rsids1.Append(rsid22);
            rsids1.Append(rsid23);
            rsids1.Append(rsid24);
            rsids1.Append(rsid25);
            rsids1.Append(rsid26);
            rsids1.Append(rsid27);
            rsids1.Append(rsid28);
            rsids1.Append(rsid29);
            rsids1.Append(rsid30);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };
            DoNotIncludeSubdocsInStats doNotIncludeSubdocsInStats1 = new DoNotIncludeSubdocsInStats();

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };

            settings1.Append(zoom1);
            settings1.Append(embedSystemFonts1);
            settings1.Append(proofState1);
            settings1.Append(stylePaneFormatFilter1);
            settings1.Append(defaultTabStop1);
            settings1.Append(hyphenationZone1);
            settings1.Append(doNotHyphenateCaps1);
            settings1.Append(drawingGridHorizontalSpacing1);
            settings1.Append(drawingGridVerticalSpacing1);
            settings1.Append(displayHorizontalDrawingGrid1);
            settings1.Append(displayVerticalDrawingGrid1);
            settings1.Append(doNotUseMarginsForDrawingGridOrigin1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(doNotValidateAgainstSchema1);
            settings1.Append(doNotDemarcateInvalidXml1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(doNotIncludeSubdocsInStats1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "cisco";
            document.PackageProperties.Title = "Приложение No";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-01-08T18:07:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-01-08T18:07:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Boyanov";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2013-04-02T06:45:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }


    }
}

