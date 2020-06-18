__all__ = ['OpenFormat',
           'MsoEncoding',
           'DocumentDirection',
           'SaveOptions',
           'OriginalFormat',
           'SaveFormat',
           'LineEndingType',
           'CompatibilityMode',
           'EditorType',
           'StoryType',
           'TableFormat',
           'TableDirection',
           'TableFieldSeparator',
           'RowHeightRule',
           'RulerStyle']


class RulerStyle:
    AdjustFirstColumn = 2  # Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.
    AdjustNone = 0  # Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.
    AdjustProportional = 1  # Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.
    AdjustSameWidth = 3  # Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.


class RowHeightRule:
    RowHeightAtLeast = 1  # The row height is at least a minimum specified value.
    RowHeightAuto = 0  # The row height is adjusted to accommodate the tallest value in the row.
    RowHeightExactly = 2  # The row height is an exact value.


class TableFieldSeparator:
    SeparateByCommas = 2  # A comma.
    SeparateByDefaultListSeparator = 3  # The default list separator.
    SeparateByParagraphs = 0  # Paragraph markers.
    SeparateByTabs = 1  # A tab.


class TableDirection:
    TableDirectionLtr = 1  # The selected rows are arranged with the first column in the leftmost position.
    TableDirectionRtl = 0  # The selected rows are arranged with the first column in the rightmost position.


class TableFormat:
    TableFormat3DEffects1 = 32  # 3D effects format number 1.
    TableFormat3DEffects2 = 33  # 3D effects format number 2.
    TableFormat3DEffects3 = 34  # 3D effects format number 3.
    TableFormatClassic1 = 4  # Classic format number 1.
    TableFormatClassic2 = 5  # Classic format number 2.
    TableFormatClassic3 = 6  # Classic format number 3.
    TableFormatClassic4 = 7  # Classic format number 4.
    TableFormatColorful1 = 8  # Colorful format number 1.
    TableFormatColorful2 = 9  # Colorful format number 2.
    TableFormatColorful3 = 10  # Colorful format number 3.
    TableFormatColumns1 = 11  # Columns format number 1.
    TableFormatColumns2 = 12  # Columns format number 2.
    TableFormatColumns3 = 13  # Columns format number 3.
    TableFormatColumns4 = 14  # Columns format number 4.
    TableFormatColumns5 = 15  # Columns format number 5.
    TableFormatContemporary = 35  # Contemporary format.
    TableFormatElegant = 36  # Elegant format.
    TableFormatGrid1 = 16  # Grid format number 1.
    TableFormatGrid2 = 17  # Grid format number 2.
    TableFormatGrid3 = 18  # Grid format number 3.
    TableFormatGrid4 = 19  # Grid format number 4.
    TableFormatGrid5 = 20  # Grid format number 5.
    TableFormatGrid6 = 21  # Grid format number 6.
    TableFormatGrid7 = 22  # Grid format number 7.
    TableFormatGrid8 = 23  # Grid format number 8.
    TableFormatList1 = 24  # List format number 1.
    TableFormatList2 = 25  # List format number 2.
    TableFormatList3 = 26  # List format number 3.
    TableFormatList4 = 27  # List format number 4.
    TableFormatList5 = 28  # List format number 5.
    TableFormatList6 = 29  # List format number 6.
    TableFormatList7 = 30  # List format number 7.
    TableFormatList8 = 31  # List format number 8.
    TableFormatNone = 0  # No formatting.
    TableFormatProfessional = 37  # Professional format.
    TableFormatSimple1 = 1  # Simple format number 1.
    TableFormatSimple2 = 2  # Simple format number 2.
    TableFormatSimple3 = 3  # Simple format number 3.
    TableFormatSubtle1 = 38  # Subtle format number 1.
    TableFormatSubtle2 = 39  # Subtle format number 2.
    TableFormatWeb1 = 40  # Web format number 1.
    TableFormatWeb2 = 41  # Web format number 2.
    TableFormatWeb3 = 42  # Web format number 3.


class OpenFormat:
    OpenFormatAllWord = 6  # A Microsoft Word format that is backward compatible with earlier versions of Word.
    OpenFormatAuto = 0  # The existing format.
    OpenFormatDocument = 1  # Word format.
    OpenFormatEncodedText = 5  # Encoded text format.
    OpenFormatRTF = 3  # Rich text format (RTF).
    OpenFormatTemplate = 2  # As a Word template.
    OpenFormatText = 4  # Unencoded text format.
    OpenFormatOpenDocumentText = 18  # OpenDocument Text format.
    OpenFormatUnicodeText = 5  # Unicode text format.
    OpenFormatWebPages = 7  # HTML format.
    OpenFormatXML = 8  # XML format.
    OpenFormatAllWordTemplates = 13  # Word template format.
    OpenFormatDocument97 = 1  # Microsoft Word 97 document format.
    OpenFormatTemplate97 = 2  # Word 97 template format.
    OpenFormatXMLDocument = 9  # XML document format.
    OpenFormatXMLDocumentSerialized = 14  # Open XML file format saved as a single XML file.
    OpenFormatXMLDocumentMacroEnabled = 10  # XML document format with macros enabled.
    OpenFormatXMLDocumentMacroEnabledSerialized = 15  # Open XML file format with macros enabled saved as a single XML file.
    OpenFormatXMLTemplate = 11  # XML template format.
    OpenFormatXMLTemplateSerialized = 16  # Open XML template format saved as a XML single file.
    OpenFormatXMLTemplateMacroEnabled = 12  # XML template format with macros enabled.
    OpenFormatXMLTemplateMacroEnabledSerialized = 17  # Open XML template format with macros enabled saved as a single XML file.


class MsoEncoding:
    EncodingArabic = 1256  # Arabic
    EncodingArabicASMO = 708  # Arabic ASMO
    EncodingArabicAutoDetect = 51256  # Web browser auto-detects type of Arabic encoding to use.
    EncodingArabicTransparentASMO = 720  # Transparent Arabic
    EncodingAutoDetect = 50001  # Web browser auto-detects type of encoding to use.
    EncodingBaltic = 1257  # Baltic
    EncodingCentralEuropean = 1250  # Central European
    EncodingCyrillic = 1251  # Cyrillic
    EncodingCyrillicAutoDetect = 51251  # Web browser auto-detects type of Cyrillic encoding to use.
    EncodingEBCDICArabic = 20420  # Extended Binary Coded Decimal Interchange Code (EBCDIC) Arabic
    EncodingEBCDICDenmarkNorway = 20277  # EBCDIC as used in Denmark and Norway
    EncodingEBCDICFinlandSweden = 20278  # EBCDIC as used in Finland and Sweden
    EncodingEBCDICFrance = 20297  # EBCDIC as used in France
    EncodingEBCDICGermany = 20273  # EBCDIC as used in Germany
    EncodingEBCDICGreek = 20423  # EBCDIC as used in the Greek language
    EncodingEBCDICGreekModern = 875  # EBCDIC as used in the Modern Greek language
    EncodingEBCDICHebrew = 20424  # EBCDIC as used in the Hebrew language
    EncodingEBCDICIcelandic = 20871  # EBCDIC as used in Iceland
    EncodingEBCDICInternational = 500  # International EBCDIC
    EncodingEBCDICItaly = 20280  # EBCDIC as used in Italy
    EncodingEBCDICJapaneseKatakanaExtended = 20290  # EBCDIC as used with Japanese Katakana (extended)
    EncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930  # EBCDIC as used with Japanese Katakana (extended) and Japanese
    EncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939  # EBCDIC as used with Japanese Latin (extended) and Japanese
    EncodingEBCDICKoreanExtended = 20833  # EBCDIC as used with Korean (extended)
    EncodingEBCDICKoreanExtendedAndKorean = 50933  # EBCDIC as used with Korean (extended) and Korean
    EncodingEBCDICLatinAmericaSpain = 20284  # EBCDIC as used in Latin America and Spain
    EncodingEBCDICMultilingualROECELatin2 = 870  # EBCDIC Multilingual ROECE (Latin 2)
    EncodingEBCDICRussian = 20880  # EBCDIC as used with Russian
    EncodingEBCDICSerbianBulgarian = 21025  # EBCDIC as used with Serbian and Bulgarian
    EncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935  # EBCDIC as used with Simplified Chinese (extended) and Simplified Chinese
    EncodingEBCDICThai = 20838  # EBCDIC as used with Thai
    EncodingEBCDICTurkish = 20905  # EBCDIC as used with Turkish
    EncodingEBCDICTurkishLatin5 = 1026  # EBCDIC as used with Turkish (Latin 5)
    EncodingEBCDICUnitedKingdom = 20285  # EBCDIC as used in the United Kingdom
    EncodingEBCDICUSCanada = 37  # EBCDIC as used in the United States and Canada
    EncodingEBCDICUSCanadaAndJapanese = 50931  # EBCDIC as used in the United States and Canada, and with Japanese
    EncodingEBCDICUSCanadaAndTraditionalChinese = 50937  # EBCDIC as used in the United States and Canada, and with Traditional Chinese
    EncodingEUCChineseSimplifiedChinese = 51936  # Extended Unix Code (EUC) as used with Chinese and Simplified Chinese
    EncodingEUCJapanese = 51932  # EUC as used with Japanese
    EncodingEUCKorean = 51949  # EUC as used with Korean
    EncodingEUCTaiwaneseTraditionalChinese = 51950  # EUC as used with Taiwanese and Traditional Chinese
    EncodingEuropa3 = 29001  # Europa
    EncodingExtAlphaLowercase = 21027  # Extended Alpha lowercase
    EncodingGreek = 1253  # Greek
    EncodingGreekAutoDetect = 51253  # Web browser auto-detects type of Greek encoding to use.
    EncodingHebrew = 1255  # Hebrew
    EncodingHZGBSimplifiedChinese = 52936  # Simplified Chinese (HZGB)
    EncodingIA5German = 20106  # German (International Alphabet No. 5, or IA5)
    EncodingIA5IRV = 20105  # IA5, International Reference Version (IRV)
    EncodingIA5Norwegian = 20108  # IA5 as used with Norwegian
    EncodingIA5Swedish = 20107  # IA5 as used with Swedish
    EncodingISCIIAssamese = 57006  # Indian Script Code for Information Interchange (ISCII) as used with Assamese
    EncodingISCIIBengali = 57003  # ISCII as used with Bengali
    EncodingISCIIDevanagari = 57002  # ISCII as used with Devanagari
    EncodingISCIIGujarati = 57010  # ISCII as used with Gujarati
    EncodingISCIIKannada = 57008  # ISCII as used with Kannada
    EncodingISCIIMalayalam = 57009  # ISCII as used with Malayalam
    EncodingISCIIOriya = 57007  # ISCII as used with Oriya
    EncodingISCIIPunjabi = 57011  # ISCII as used with Punjabi
    EncodingISCIITamil = 57004  # ISCII as used with Tamil
    EncodingISCIITelugu = 57005  # ISCII as used with Telugu
    EncodingISO2022CNSimplifiedChinese = 50229  # ISO 2022-CN encoding as used with Simplified Chinese
    EncodingISO2022CNTraditionalChinese = 50227  # ISO 2022-CN encoding as used with Traditional Chinese
    EncodingISO2022JPJISX02011989 = 50222  # ISO 2022-JP
    EncodingISO2022JPJISX02021984 = 50221  # ISO 2022-JP
    EncodingISO2022JPNoHalfwidthKatakana = 50220  # ISO 2022-JP with no half-width Katakana
    EncodingISO2022KR = 50225  # ISO 2022-KR
    EncodingISO6937NonSpacingAccent = 20269  # ISO 6937 Non-Spacing Accent
    EncodingISO885915Latin9 = 28605  # ISO 8859-15 with Latin 9
    EncodingISO88591Latin1 = 28591  # ISO 8859-1 Latin 1
    EncodingISO88592CentralEurope = 28592  # ISO 8859-2 Central Europe
    EncodingISO88593Latin3 = 28593  # ISO 8859-3 Latin 3
    EncodingISO88594Baltic = 28594  # ISO 8859-4 Baltic
    EncodingISO88595Cyrillic = 28595  # ISO 8859-5 Cyrillic
    EncodingISO88596Arabic = 28596  # ISA 8859-6 Arabic
    EncodingISO88597Greek = 28597  # ISO 8859-7 Greek
    EncodingISO88598Hebrew = 28598  # ISO 8859-8 Hebrew
    EncodingISO88598HebrewLogical = 38598  # ISO 8859-8 Hebrew (Logical)
    EncodingISO88599Turkish = 28599  # ISO 8859-9 Turkish
    EncodingJapaneseAutoDetect = 50932  # Web browser auto-detects type of Japanese encoding to use.
    EncodingJapaneseShiftJIS = 932  # Japanese (Shift-JIS)
    EncodingKOI8R = 20866  # KOI8-R
    EncodingKOI8U = 21866  # K0I8-U
    EncodingKorean = 949  # Korean
    EncodingKoreanAutoDetect = 50949  # Web browser auto-detects type of Korean encoding to use.
    EncodingKoreanJohab = 1361  # Korean (Johab)
    EncodingMacArabic = 10004  # Macintosh Arabic
    EncodingMacCroatia = 10082  # Macintosh Croatian
    EncodingMacCyrillic = 10007  # Macintosh Cyrillic
    EncodingMacGreek1 = 10006  # Macintosh Greek
    EncodingMacHebrew = 10005  # Macintosh Hebrew
    EncodingMacIcelandic = 10079  # Macintosh Icelandic
    EncodingMacJapanese = 10001  # Macintosh Japanese
    EncodingMacKorean = 10003  # Macintosh Korean
    EncodingMacLatin2 = 10029  # Macintosh Latin 2
    EncodingMacRoman = 10000  # Macintosh Roman
    EncodingMacRomania = 10010  # Macintosh Romanian
    EncodingMacSimplifiedChineseGB2312 = 10008  # Macintosh Simplified Chinese (GB 2312)
    EncodingMacTraditionalChineseBig5 = 10002  # Macintosh Traditional Chinese (Big 5)
    EncodingMacTurkish = 10081  # Macintosh Turkish
    EncodingMacUkraine = 10017  # Macintosh Ukrainian
    EncodingOEMArabic = 864  # OEM as used with Arabic
    EncodingOEMBaltic = 775  # OEM as used with Baltic
    EncodingOEMCanadianFrench = 863  # OEM as used with Canadian French
    EncodingOEMCyrillic = 855  # OEM as used with Cyrillic
    EncodingOEMCyrillicII = 866  # OEM as used with Cyrillic II
    EncodingOEMGreek437G = 737  # OEM as used with Greek 437G
    EncodingOEMHebrew = 862  # OEM as used with Hebrew
    EncodingOEMIcelandic = 861  # OEM as used with Icelandic
    EncodingOEMModernGreek = 869  # OEM as used with Modern Greek
    EncodingOEMMultilingualLatinI = 850  # OEM as used with multi-lingual Latin I
    EncodingOEMMultilingualLatinII = 852  # OEM as used with multi-lingual Latin II
    EncodingOEMNordic = 865  # OEM as used with Nordic languages
    EncodingOEMPortuguese = 860  # OEM as used with Portuguese
    EncodingOEMTurkish = 857  # OEM as used with Turkish
    EncodingOEMUnitedStates = 437  # OEM as used in the United States
    EncodingSimplifiedChineseAutoDetect = 50936  # Web browser auto-detects type of Simplified Chinese encoding to use.
    EncodingSimplifiedChineseGB18030 = 54936  # Simplified Chinese GB 18030
    EncodingSimplifiedChineseGBK = 936  # Simplified Chinese GBK
    EncodingT61 = 20261  # T61
    EncodingTaiwanCNS = 20000  # Taiwan CNS
    EncodingTaiwanEten = 20002  # Taiwan Eten
    EncodingTaiwanIBM5550 = 20003  # Taiwan IBM 5550
    EncodingTaiwanTCA = 20001  # Taiwan TCA
    EncodingTaiwanTeleText = 20004  # Taiwan Teletext
    EncodingTaiwanWang = 20005  # Taiwan Wang
    EncodingThai = 874  # Thai
    EncodingTraditionalChineseAutoDetect = 50950  # Web browser auto-detects type of Traditional Chinese encoding to use.
    EncodingTraditionalChineseBig5 = 950  # Traditional Chinese Big 5
    EncodingTurkish = 1254  # Turkish
    EncodingUnicodeBigEndian = 1201  # Unicode big endian
    EncodingUnicodeLittleEndian = 1200  # Unicode little endian
    EncodingUSASCII = 20127  # United States ASCII
    EncodingUTF7 = 65000  # UTF-7 encoding
    EncodingUTF8 = 65001  # UTF-8 encoding
    EncodingVietnamese = 1258  # Vietnamese
    EncodingWestern = 1252  # Western


class DocumentDirection:
    LeftToRight = 0
    RightToLeft = 1


class SaveOptions:
    DoNotSaveChanges = 0  # Do not save pending changes.
    PromptToSaveChanges = -2  # Prompt the user to save pending changes.
    SaveChanges = -1  # Save pending changes automatically without prompting the user.


class OriginalFormat:
    OriginalDocumentFormat = 1  # Original document format.
    PromptUser = 2  # Prompt user to select a document format.
    WordDocument = 0  # Microsoft Word document format.


class SaveFormat:
    FormatDocument = 0  # Microsoft Office Word 97 - 2003 binary file format.
    FormatDOSText = 4  # Microsoft DOS text format.
    FormatDOSTextLineBreaks = 5  # Microsoft DOS text with line breaks preserved.
    FormatEncodedText = 7  # Encoded text format.
    FormatFilteredHTML = 10  # Filtered HTML format.
    FormatFlatXML = 19  # Open XML file format saved as a single XML file.
    FormatFlatXMLMacroEnabled = 20  # Open XML file format with macros enabled saved as a single XML file.
    FormatFlatXMLTemplate = 21  # Open XML template format saved as a XML single file.
    FormatFlatXMLTemplateMacroEnabled = 22  # Open XML template format with macros enabled saved as a single XML file.
    FormatOpenDocumentText = 23  # OpenDocument Text format.
    FormatHTML = 8  # Standard HTML format.
    FormatRTF = 6  # Rich text format (RTF).
    FormatStrictOpenXMLDocument = 24  # Strict Open XML document format.
    FormatTemplate = 1  # Word template format.
    FormatText = 2  # Microsoft Windows text format.
    FormatTextLineBreaks = 3  # Windows text format with line breaks preserved.
    FormatUnicodeText = 7  # Unicode text format.
    FormatWebArchive = 9  # Web archive format.
    FormatXML = 11  # Extensible Markup Language (XML) format.
    FormatDocument97 = 0  # Microsoft Word 97 document format.
    FormatDocumentDefault = 16  # Word default document file format. For Word, this is the DOCX format.
    FormatPDF = 17  # PDF format.
    FormatTemplate97 = 1  # Word 97 template format.
    FormatXMLDocument = 12  # XML document format.
    FormatXMLDocumentMacroEnabled = 13  # XML document format with macros enabled.
    FormatXMLTemplate = 14  # XML template format.
    FormatXMLTemplateMacroEnabled = 15  # XML template format with macros enabled.
    FormatXPS = 18  # XPS format.


class LineEndingType:
    CRLF = 0  # Carriage return plus line feed.
    CROnly = 1  # Carriage return only.
    LFCR = 3  # Line feed plus carriage return.
    LFOnly = 2  # Line feed only.
    LSPS = 4  # Not supported.


class CompatibilityMode:
    Current = 65535  # Compatibility mode equivalent to the latest version of Word.
    Word2003 = 11  # Word is put into a mode that is most compatible with Word 2003. Features new to Word are disabled in this mode.
    Word2007 = 12  # Word is put into a mode that is most compatible with Word 2007. Features new to Word are disabled in this mode.
    Word2010 = 14  # Word is put into a mode that is most compatible with Word 2010. Features new to Word are disabled in this mode.
    Word2013 = 15  # Default. All Word features are enabled.


class EditorType:
    EditorCurrent = -6  # Represents the current user of the document.
    EditorEditors = -5  # Represents the Editors group for documents that use Information Rights Management.
    EditorEveryone = -1  # Represents all users who open a document.
    EditorOwners = -4  # Represents the Owners group for documents that use Information Rights Management.


class StoryType:
    CommentsStory = 4  # Comments story.
    EndnoteContinuationNoticeStory = 17  # Endnote continuation notice story.
    EndnoteContinuationSeparatorStory = 16  # Endnote continuation separator story.
    EndnoteSeparatorStory = 15  # Endnote separator story.
    EndnotesStory = 3  # Endnotes story.
    EvenPagesFooterStory = 8  # Even pages footer story.
    EvenPagesHeaderStory = 6  # Even pages header story.
    FirstPageFooterStory = 11  # First page footer story.
    FirstPageHeaderStory = 10  # First page header story.
    FootnoteContinuationNoticeStory = 14  # Footnote continuation notice story.
    FootnoteContinuationSeparatorStory = 13  # Footnote continuation separator story.
    FootnoteSeparatorStory = 12  # Footnote separator story.
    FootnotesStory = 2  # Footnotes story.
    MainTextStory = 1  # Main text story.
    PrimaryFooterStory = 9  # Primary footer story.
    PrimaryHeaderStory = 7  # Primary header story.
    TextFrameStory = 5  # Text frame story.
