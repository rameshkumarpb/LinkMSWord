/*
 * MicrosoftWord.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class MicrosoftWordBaseObject, MicrosoftWordBaseApplication, MicrosoftWordBaseDocument, MicrosoftWordBasicWindow, MicrosoftWordPrintSettings, MicrosoftWordCommandBarControl, MicrosoftWordCommandBarButton, MicrosoftWordCommandBarCombobox, MicrosoftWordCommandBarPopup, MicrosoftWordCommandBar, MicrosoftWordDocumentProperty, MicrosoftWordCustomDocumentProperty, MicrosoftWordWebPageFont, MicrosoftWordWordComment, MicrosoftWordWordList, MicrosoftWordWordOptions, MicrosoftWordAddIn, MicrosoftWordApplication, MicrosoftWordAutoTextEntry, MicrosoftWordBookmark, MicrosoftWordBorderOptions, MicrosoftWordBorder, MicrosoftWordBrowser, MicrosoftWordBuildingBlockCategory, MicrosoftWordBuildingBlockType, MicrosoftWordBuildingBlock, MicrosoftWordCaptionLabel, MicrosoftWordCheckBox, MicrosoftWordCoauthLock, MicrosoftWordCoauthUpdate, MicrosoftWordCoauthor, MicrosoftWordCoauthoring, MicrosoftWordConflict, MicrosoftWordCustomLabel, MicrosoftWordDataMergeDataField, MicrosoftWordDataMergeDataSource, MicrosoftWordDataMergeFieldName, MicrosoftWordDataMergeField, MicrosoftWordDataMerge, MicrosoftWordDefaultWebOptions, MicrosoftWordDialog, MicrosoftWordDocumentVersion, MicrosoftWordDocument, MicrosoftWordDropCap, MicrosoftWordDropDown, MicrosoftWordEndnoteOptions, MicrosoftWordEndnote, MicrosoftWordEnvelope, MicrosoftWordFieldOptions, MicrosoftWordField, MicrosoftWordFileConverter, MicrosoftWordFind, MicrosoftWordFont, MicrosoftWordFootnoteOptions, MicrosoftWordFootnote, MicrosoftWordFormField, MicrosoftWordFrame, MicrosoftWordHeaderFooter, MicrosoftWordHeadingStyle, MicrosoftWordHyperlinkObject, MicrosoftWordIndex, MicrosoftWordKeyBinding, MicrosoftWordLetterContent, MicrosoftWordLineNumbering, MicrosoftWordLinkFormat, MicrosoftWordListEntry, MicrosoftWordListFormat, MicrosoftWordListGallery, MicrosoftWordListLevel, MicrosoftWordListTemplate, MicrosoftWordMailingLabel, MicrosoftWordMathAccent, MicrosoftWordMathAutocorrectEntry, MicrosoftWordMathAutocorrect, MicrosoftWordMathBar, MicrosoftWordMathBorderBox, MicrosoftWordMathBox, MicrosoftWordMathBreak, MicrosoftWordMathDelimiter, MicrosoftWordMathEquationArray, MicrosoftWordMathFraction, MicrosoftWordMathFunc, MicrosoftWordMathFunction, MicrosoftWordMathGroupChar, MicrosoftWordMathLeftScripts, MicrosoftWordMathLowerLimit, MicrosoftWordMathMatrixColumn, MicrosoftWordMathMatrixRow, MicrosoftWordMathMatrix, MicrosoftWordMathNary, MicrosoftWordMathObject, MicrosoftWordMathPhantom, MicrosoftWordMathRadical, MicrosoftWordMathRecognizedFunction, MicrosoftWordMathSubAndSuperScript, MicrosoftWordMathSubscript, MicrosoftWordMathSuperscript, MicrosoftWordMathUpperLimit, MicrosoftWordPageNumberOptions, MicrosoftWordPageNumber, MicrosoftWordPageSetup, MicrosoftWordPane, MicrosoftWordRangeEndnoteOptions, MicrosoftWordRangeFootnoteOptions, MicrosoftWordRecentFile, MicrosoftWordReplacement, MicrosoftWordReviewer, MicrosoftWordRevision, MicrosoftWordSelectionObject, MicrosoftWordSubdocument, MicrosoftWordSystemObject, MicrosoftWordTabStop, MicrosoftWordTableOfAuthorities, MicrosoftWordTableOfContents, MicrosoftWordTableOfFigures, MicrosoftWordTemplate, MicrosoftWordTextColumn, MicrosoftWordTextInput, MicrosoftWordTextRetrievalMode, MicrosoftWordVariable, MicrosoftWordView, MicrosoftWordWebOptions, MicrosoftWordWindow, MicrosoftWordZoom, MicrosoftWordAdjustment, MicrosoftWordCalloutFormat, MicrosoftWordFillFormat, MicrosoftWordGlowFormat, MicrosoftWordHorizontalLineFormat, MicrosoftWordInlineShape, MicrosoftWordInlineHorizontalLine, MicrosoftWordInlinePictureBullet, MicrosoftWordInlinePicture, MicrosoftWordLineFormat, MicrosoftWordOfficeTheme, MicrosoftWordPictureFormat, MicrosoftWordReflectionFormat, MicrosoftWordShadowFormat, MicrosoftWordShape, MicrosoftWordCallout, MicrosoftWordLineShape, MicrosoftWordPicture, MicrosoftWordSoftEdgeFormat, MicrosoftWordStandardInlineHorizontalLine, MicrosoftWordTextBox, MicrosoftWordTextFrame, MicrosoftWordThemeColorScheme, MicrosoftWordThemeColor, MicrosoftWordThemeEffectScheme, MicrosoftWordThemeFontScheme, MicrosoftWordThemeFont, MicrosoftWordMajorThemeFont, MicrosoftWordMinorThemeFont, MicrosoftWordThemeFonts, MicrosoftWordThreeDFormat, MicrosoftWordWordArtFormat, MicrosoftWordWordArt, MicrosoftWordWrapFormat, MicrosoftWordWordStyle, MicrosoftWordParagraphFormat, MicrosoftWordParagraph, MicrosoftWordSection, MicrosoftWordShading, MicrosoftWordTextRange, MicrosoftWordCharacter, MicrosoftWordGrammaticalError, MicrosoftWordSentence, MicrosoftWordSpellingError, MicrosoftWordStoryRange, MicrosoftWordWord, MicrosoftWordAutocorrectEntry, MicrosoftWordAutocorrect, MicrosoftWordDictionary, MicrosoftWordFirstLetterException, MicrosoftWordLanguage, MicrosoftWordOtherCorrectionsException, MicrosoftWordReadabilityStatistic, MicrosoftWordSynonymInfo, MicrosoftWordTwoInitialCapsException, MicrosoftWordCell, MicrosoftWordColumnOptions, MicrosoftWordColumn, MicrosoftWordCondition, MicrosoftWordRowOptions, MicrosoftWordRow, MicrosoftWordTableStyle, MicrosoftWordTable;

enum MicrosoftWordSavo {
	MicrosoftWordSavoYes = 'yes ' /* Save objects now */,
	MicrosoftWordSavoNo = 'no  ' /* Do not save objects */,
	MicrosoftWordSavoAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum MicrosoftWordSavo MicrosoftWordSavo;

enum MicrosoftWordKfrm {
	MicrosoftWordKfrmIndex = 'indx' /* keyform designating indexed access */,
	MicrosoftWordKfrmNamed = 'name' /* keyform designating named access */,
	MicrosoftWordKfrmId = 'ID  ' /* keyform designating access by unique identifier */
};
typedef enum MicrosoftWordKfrm MicrosoftWordKfrm;

enum MicrosoftWordEnum {
	MicrosoftWordEnumStandard = 'lwst' /* Standard PostScript error handling   */,
	MicrosoftWordEnumDetailed = 'lwdt' /* print a detailed report of Postscript errors */
};
typedef enum MicrosoftWordEnum MicrosoftWordEnum;

enum MicrosoftWordMlDs {
	MicrosoftWordMlDsLineDashStyleUnset = '\000\222\377\376',
	MicrosoftWordMlDsLineDashStyleSolid = '\000\223\000\001',
	MicrosoftWordMlDsLineDashStyleSquareDot = '\000\223\000\002',
	MicrosoftWordMlDsLineDashStyleRoundDot = '\000\223\000\003',
	MicrosoftWordMlDsLineDashStyleDash = '\000\223\000\004',
	MicrosoftWordMlDsLineDashStyleDashDot = '\000\223\000\005',
	MicrosoftWordMlDsLineDashStyleDashDotDot = '\000\223\000\006',
	MicrosoftWordMlDsLineDashStyleLongDash = '\000\223\000\007',
	MicrosoftWordMlDsLineDashStyleLongDashDot = '\000\223\000\010',
	MicrosoftWordMlDsLineDashStyleLongDashDotDot = '\000\223\000\011',
	MicrosoftWordMlDsLineDashStyleSystemDash = '\000\223\000\012',
	MicrosoftWordMlDsLineDashStyleSystemDot = '\000\223\000\013',
	MicrosoftWordMlDsLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum MicrosoftWordMlDs MicrosoftWordMlDs;

enum MicrosoftWordMLnS {
	MicrosoftWordMLnSLineStyleUnset = '\000\224\377\376',
	MicrosoftWordMLnSSingleLine = '\000\225\000\001',
	MicrosoftWordMLnSThinThinLine = '\000\225\000\002',
	MicrosoftWordMLnSThinThickLine = '\000\225\000\003',
	MicrosoftWordMLnSThickThinLine = '\000\225\000\004',
	MicrosoftWordMLnSThickBetweenThinLine = '\000\225\000\005'
};
typedef enum MicrosoftWordMLnS MicrosoftWordMLnS;

enum MicrosoftWordMAhS {
	MicrosoftWordMAhSArrowheadStyleUnset = '\000\221\377\376',
	MicrosoftWordMAhSNoArrowhead = '\000\222\000\001',
	MicrosoftWordMAhSTriangleArrowhead = '\000\222\000\002',
	MicrosoftWordMAhSOpen_arrowhead = '\000\222\000\003',
	MicrosoftWordMAhSStealthArrowhead = '\000\222\000\004',
	MicrosoftWordMAhSDiamondArrowhead = '\000\222\000\005',
	MicrosoftWordMAhSOvalArrowhead = '\000\222\000\006'
};
typedef enum MicrosoftWordMAhS MicrosoftWordMAhS;

enum MicrosoftWordMAhW {
	MicrosoftWordMAhWArrowheadWidthUnset = '\000\220\377\376',
	MicrosoftWordMAhWNarrowWidthArrowhead = '\000\221\000\001',
	MicrosoftWordMAhWMediumWidthArrowhead = '\000\221\000\002',
	MicrosoftWordMAhWWideArrowhead = '\000\221\000\003'
};
typedef enum MicrosoftWordMAhW MicrosoftWordMAhW;

enum MicrosoftWordMAhL {
	MicrosoftWordMAhLArrowheadLengthUnset = '\000\223\377\376',
	MicrosoftWordMAhLShortArrowhead = '\000\224\000\001',
	MicrosoftWordMAhLMediumArrowhead = '\000\224\000\002',
	MicrosoftWordMAhLLongArrowhead = '\000\224\000\003'
};
typedef enum MicrosoftWordMAhL MicrosoftWordMAhL;

enum MicrosoftWordMFdT {
	MicrosoftWordMFdTFillUnset = '\000c\377\376',
	MicrosoftWordMFdTFillSolid = '\000d\000\001',
	MicrosoftWordMFdTFillPatterned = '\000d\000\002',
	MicrosoftWordMFdTFillGradient = '\000d\000\003',
	MicrosoftWordMFdTFillTextured = '\000d\000\004',
	MicrosoftWordMFdTFillBackground = '\000d\000\005',
	MicrosoftWordMFdTFillPicture = '\000d\000\006'
};
typedef enum MicrosoftWordMFdT MicrosoftWordMFdT;

enum MicrosoftWordMGdS {
	MicrosoftWordMGdSGradientUnset = '\000d\377\376',
	MicrosoftWordMGdSHorizontalGradient = '\000e\000\001',
	MicrosoftWordMGdSVerticalGradient = '\000e\000\002',
	MicrosoftWordMGdSDiagonalUpGradient = '\000e\000\003',
	MicrosoftWordMGdSDiagonalDownGradient = '\000e\000\004',
	MicrosoftWordMGdSFromCornerGradient = '\000e\000\005',
	MicrosoftWordMGdSFromTitleGradient = '\000e\000\006',
	MicrosoftWordMGdSFromCenterGradient = '\000e\000\007'
};
typedef enum MicrosoftWordMGdS MicrosoftWordMGdS;

enum MicrosoftWordMGCt {
	MicrosoftWordMGCtGradientTypeUnset = '\003\357\377\376',
	MicrosoftWordMGCtSingleShadeGradientType = '\003\360\000\001',
	MicrosoftWordMGCtTwoColorsGradientType = '\003\360\000\002',
	MicrosoftWordMGCtPresetColorsGradientType = '\003\360\000\003',
	MicrosoftWordMGCtMultiColorsGradientType = '\003\360\000\004'
};
typedef enum MicrosoftWordMGCt MicrosoftWordMGCt;

enum MicrosoftWordMxtT {
	MicrosoftWordMxtTTextureTypeTextureTypeUnset = '\003\360\377\376',
	MicrosoftWordMxtTTextureTypePresetTexture = '\003\361\000\001',
	MicrosoftWordMxtTTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum MicrosoftWordMxtT MicrosoftWordMxtT;

enum MicrosoftWordMPzT {
	MicrosoftWordMPzTPresetTextureUnset = '\000e\377\376',
	MicrosoftWordMPzTTexturePapyrus = '\000f\000\001',
	MicrosoftWordMPzTTextureCanvas = '\000f\000\002',
	MicrosoftWordMPzTTextureDenim = '\000f\000\003',
	MicrosoftWordMPzTTextureWovenMat = '\000f\000\004',
	MicrosoftWordMPzTTextureWaterDroplets = '\000f\000\005',
	MicrosoftWordMPzTTexturePaperBag = '\000f\000\006',
	MicrosoftWordMPzTTextureFishFossil = '\000f\000\007',
	MicrosoftWordMPzTTextureSand = '\000f\000\010',
	MicrosoftWordMPzTTextureGreenMarble = '\000f\000\011',
	MicrosoftWordMPzTTextureWhiteMarble = '\000f\000\012',
	MicrosoftWordMPzTTextureBrownMarble = '\000f\000\013',
	MicrosoftWordMPzTTextureGranite = '\000f\000\014',
	MicrosoftWordMPzTTextureNewsprint = '\000f\000\015',
	MicrosoftWordMPzTTextureRecycledPaper = '\000f\000\016',
	MicrosoftWordMPzTTextureParchment = '\000f\000\017',
	MicrosoftWordMPzTTextureStationery = '\000f\000\020',
	MicrosoftWordMPzTTextureBlueTissuePaper = '\000f\000\021',
	MicrosoftWordMPzTTexturePinkTissuePaper = '\000f\000\022',
	MicrosoftWordMPzTTexturePurpleMesh = '\000f\000\023',
	MicrosoftWordMPzTTextureBouquet = '\000f\000\024',
	MicrosoftWordMPzTTextureCork = '\000f\000\025',
	MicrosoftWordMPzTTextureWalnut = '\000f\000\026',
	MicrosoftWordMPzTTextureOak = '\000f\000\027',
	MicrosoftWordMPzTTextureMediumWood = '\000f\000\030'
};
typedef enum MicrosoftWordMPzT MicrosoftWordMPzT;

enum MicrosoftWordPpTy {
	MicrosoftWordPpTyPatternUnset = '\000f\377\376',
	MicrosoftWordPpTyFivePercentPattern = '\000g\000\001',
	MicrosoftWordPpTyTenPercentPattern = '\000g\000\002',
	MicrosoftWordPpTyTwentyPercentPattern = '\000g\000\003',
	MicrosoftWordPpTyTwentyFivePercentPattern = '\000g\000\004',
	MicrosoftWordPpTyThirtyPercentPattern = '\000g\000\005',
	MicrosoftWordPpTyFortyPercentPattern = '\000g\000\006',
	MicrosoftWordPpTyFiftyPercentPattern = '\000g\000\007',
	MicrosoftWordPpTySixtyPercentPattern = '\000g\000\010',
	MicrosoftWordPpTySeventyPercentPattern = '\000g\000\011',
	MicrosoftWordPpTySeventyFivePercentPattern = '\000g\000\012',
	MicrosoftWordPpTyEightyPercentPattern = '\000g\000\013',
	MicrosoftWordPpTyNinetyPercentPattern = '\000g\000\014',
	MicrosoftWordPpTyDarkHorizontalPattern = '\000g\000\015',
	MicrosoftWordPpTyDarkVerticalPattern = '\000g\000\016',
	MicrosoftWordPpTyDarkDownwardDiagonalPattern = '\000g\000\017',
	MicrosoftWordPpTyDarkUpwardDiagonalPattern = '\000g\000\020',
	MicrosoftWordPpTySmallCheckerBoardPattern = '\000g\000\021',
	MicrosoftWordPpTyTrellisPattern = '\000g\000\022',
	MicrosoftWordPpTyLightHorizontalPattern = '\000g\000\023',
	MicrosoftWordPpTyLightVerticalPattern = '\000g\000\024',
	MicrosoftWordPpTyLightDownwardDiagonalPattern = '\000g\000\025',
	MicrosoftWordPpTyLightUpwardDiagonalPattern = '\000g\000\026',
	MicrosoftWordPpTySmallGridPattern = '\000g\000\027',
	MicrosoftWordPpTyDottedDiamondPattern = '\000g\000\030',
	MicrosoftWordPpTyWideDownwardDiagonal = '\000g\000\031',
	MicrosoftWordPpTyWideUpwardDiagonalPattern = '\000g\000\032',
	MicrosoftWordPpTyDashedUpwardDiagonalPattern = '\000g\000\033',
	MicrosoftWordPpTyDashedDownwardDiagonalPattern = '\000g\000\034',
	MicrosoftWordPpTyNarrowVerticalPattern = '\000g\000\035',
	MicrosoftWordPpTyNarrowHorizontalPattern = '\000g\000\036',
	MicrosoftWordPpTyDashedVerticalPattern = '\000g\000\037',
	MicrosoftWordPpTyDashedHorizontalPattern = '\000g\000 ',
	MicrosoftWordPpTyLargeConfettiPattern = '\000g\000!',
	MicrosoftWordPpTyLargeGridPattern = '\000g\000\"',
	MicrosoftWordPpTyHorizontalBrickPattern = '\000g\000#',
	MicrosoftWordPpTyLargeCheckerBoardPattern = '\000g\000$',
	MicrosoftWordPpTySmallConfettiPattern = '\000g\000%',
	MicrosoftWordPpTyZigZagPattern = '\000g\000&',
	MicrosoftWordPpTySolidDiamondPattern = '\000g\000\'',
	MicrosoftWordPpTyDiagonalBrickPattern = '\000g\000(',
	MicrosoftWordPpTyOutlinedDiamondPattern = '\000g\000)',
	MicrosoftWordPpTyPlaidPattern = '\000g\000*',
	MicrosoftWordPpTySpherePattern = '\000g\000+',
	MicrosoftWordPpTyWeavePattern = '\000g\000,',
	MicrosoftWordPpTyDottedGridPattern = '\000g\000-',
	MicrosoftWordPpTyDivotPattern = '\000g\000.',
	MicrosoftWordPpTyShinglePattern = '\000g\000/',
	MicrosoftWordPpTyWavePattern = '\000g\0000',
	MicrosoftWordPpTyHorizontalPattern = '\000g\0001',
	MicrosoftWordPpTyVerticalPattern = '\000g\0002',
	MicrosoftWordPpTyCrossPattern = '\000g\0003',
	MicrosoftWordPpTyDownwardDiagonalPattern = '\000g\0004',
	MicrosoftWordPpTyUpwardDiagonalPattern = '\000g\0005',
	MicrosoftWordPpTyDiagonalCrossPattern = '\000g\0005'
};
typedef enum MicrosoftWordPpTy MicrosoftWordPpTy;

enum MicrosoftWordMPGb {
	MicrosoftWordMPGbPresetGradientUnset = '\000g\377\376',
	MicrosoftWordMPGbGradientEarlySunset = '\000h\000\001',
	MicrosoftWordMPGbGradientLateSunset = '\000h\000\002',
	MicrosoftWordMPGbGradientNightfall = '\000h\000\003',
	MicrosoftWordMPGbGradientDaybreak = '\000h\000\004',
	MicrosoftWordMPGbGradientHorizon = '\000h\000\005',
	MicrosoftWordMPGbGradientDesert = '\000h\000\006',
	MicrosoftWordMPGbGradientOcean = '\000h\000\007',
	MicrosoftWordMPGbGradientCalmWater = '\000h\000\010',
	MicrosoftWordMPGbGradientFire = '\000h\000\011',
	MicrosoftWordMPGbGradientFog = '\000h\000\012',
	MicrosoftWordMPGbGradientMoss = '\000h\000\013',
	MicrosoftWordMPGbGradientPeacock = '\000h\000\014',
	MicrosoftWordMPGbGradientWheat = '\000h\000\015',
	MicrosoftWordMPGbGradientParchment = '\000h\000\016',
	MicrosoftWordMPGbGradientMahogany = '\000h\000\017',
	MicrosoftWordMPGbGradientRainbow = '\000h\000\020',
	MicrosoftWordMPGbGradientRainbow2 = '\000h\000\021',
	MicrosoftWordMPGbGradientGold = '\000h\000\022',
	MicrosoftWordMPGbGradientGold2 = '\000h\000\023',
	MicrosoftWordMPGbGradientBrass = '\000h\000\024',
	MicrosoftWordMPGbGradientChrome = '\000h\000\025',
	MicrosoftWordMPGbGradientChrome2 = '\000h\000\026',
	MicrosoftWordMPGbGradientSilver = '\000h\000\027',
	MicrosoftWordMPGbGradientSapphire = '\000h\000\030'
};
typedef enum MicrosoftWordMPGb MicrosoftWordMPGb;

enum MicrosoftWordMSdT {
	MicrosoftWordMSdTShadowUnset = '\003_\377\376',
	MicrosoftWordMSdTShadow1 = '\003`\000\001',
	MicrosoftWordMSdTShadow2 = '\003`\000\002',
	MicrosoftWordMSdTShadow3 = '\003`\000\003',
	MicrosoftWordMSdTShadow4 = '\003`\000\004',
	MicrosoftWordMSdTShadow5 = '\003`\000\005',
	MicrosoftWordMSdTShadow6 = '\003`\000\006',
	MicrosoftWordMSdTShadow7 = '\003`\000\007',
	MicrosoftWordMSdTShadow8 = '\003`\000\010',
	MicrosoftWordMSdTShadow9 = '\003`\000\011',
	MicrosoftWordMSdTShadow10 = '\003`\000\012',
	MicrosoftWordMSdTShadow11 = '\003`\000\013',
	MicrosoftWordMSdTShadow12 = '\003`\000\014',
	MicrosoftWordMSdTShadow13 = '\003`\000\015',
	MicrosoftWordMSdTShadow14 = '\003`\000\016',
	MicrosoftWordMSdTShadow15 = '\003`\000\017',
	MicrosoftWordMSdTShadow16 = '\003`\000\020',
	MicrosoftWordMSdTShadow17 = '\003`\000\021',
	MicrosoftWordMSdTShadow18 = '\003`\000\022',
	MicrosoftWordMSdTShadow19 = '\003`\000\023',
	MicrosoftWordMSdTShadow20 = '\003`\000\024',
	MicrosoftWordMSdTShadow21 = '\003`\000\025',
	MicrosoftWordMSdTShadow22 = '\003`\000\026',
	MicrosoftWordMSdTShadow23 = '\003`\000\027',
	MicrosoftWordMSdTShadow24 = '\003`\000\030',
	MicrosoftWordMSdTShadow25 = '\003`\000\031',
	MicrosoftWordMSdTShadow26 = '\003`\000\032',
	MicrosoftWordMSdTShadow27 = '\003`\000\033',
	MicrosoftWordMSdTShadow28 = '\003`\000\034',
	MicrosoftWordMSdTShadow29 = '\003`\000\035',
	MicrosoftWordMSdTShadow30 = '\003`\000\036',
	MicrosoftWordMSdTShadow31 = '\003`\000\037',
	MicrosoftWordMSdTShadow32 = '\003`\000 ',
	MicrosoftWordMSdTShadow33 = '\003`\000!',
	MicrosoftWordMSdTShadow34 = '\003`\000\"',
	MicrosoftWordMSdTShadow35 = '\003`\000#',
	MicrosoftWordMSdTShadow36 = '\003`\000$',
	MicrosoftWordMSdTShadow37 = '\003`\000%',
	MicrosoftWordMSdTShadow38 = '\003`\000&',
	MicrosoftWordMSdTShadow39 = '\003`\000\'',
	MicrosoftWordMSdTShadow40 = '\003`\000(',
	MicrosoftWordMSdTShadow41 = '\003`\000)',
	MicrosoftWordMSdTShadow42 = '\003`\000*',
	MicrosoftWordMSdTShadow43 = '\003`\000+'
};
typedef enum MicrosoftWordMSdT MicrosoftWordMSdT;

enum MicrosoftWordMPXF {
	MicrosoftWordMPXFWordartFormatUnset = '\003\361\377\376',
	MicrosoftWordMPXFWordartFormat1 = '\003\362\000\000',
	MicrosoftWordMPXFWordartFormat2 = '\003\362\000\001',
	MicrosoftWordMPXFWordartFormat3 = '\003\362\000\002',
	MicrosoftWordMPXFWordartFormat4 = '\003\362\000\003',
	MicrosoftWordMPXFWordartFormat5 = '\003\362\000\004',
	MicrosoftWordMPXFWordartFormat6 = '\003\362\000\005',
	MicrosoftWordMPXFWordartFormat7 = '\003\362\000\006',
	MicrosoftWordMPXFWordartFormat8 = '\003\362\000\007',
	MicrosoftWordMPXFWordartFormat9 = '\003\362\000\010',
	MicrosoftWordMPXFWordartFormat10 = '\003\362\000\011',
	MicrosoftWordMPXFWordartFormat11 = '\003\362\000\012',
	MicrosoftWordMPXFWordartFormat12 = '\003\362\000\013',
	MicrosoftWordMPXFWordartFormat13 = '\003\362\000\014',
	MicrosoftWordMPXFWordartFormat14 = '\003\362\000\015',
	MicrosoftWordMPXFWordartFormat15 = '\003\362\000\016',
	MicrosoftWordMPXFWordartFormat16 = '\003\362\000\017',
	MicrosoftWordMPXFWordartFormat17 = '\003\362\000\020',
	MicrosoftWordMPXFWordartFormat18 = '\003\362\000\021',
	MicrosoftWordMPXFWordartFormat19 = '\003\362\000\022',
	MicrosoftWordMPXFWordartFormat20 = '\003\362\000\023',
	MicrosoftWordMPXFWordartFormat21 = '\003\362\000\024',
	MicrosoftWordMPXFWordartFormat22 = '\003\362\000\025',
	MicrosoftWordMPXFWordartFormat23 = '\003\362\000\026',
	MicrosoftWordMPXFWordartFormat24 = '\003\362\000\027',
	MicrosoftWordMPXFWordartFormat25 = '\003\362\000\030',
	MicrosoftWordMPXFWordartFormat26 = '\003\362\000\031',
	MicrosoftWordMPXFWordartFormat27 = '\003\362\000\032',
	MicrosoftWordMPXFWordartFormat28 = '\003\362\000\033',
	MicrosoftWordMPXFWordartFormat29 = '\003\362\000\034',
	MicrosoftWordMPXFWordartFormat30 = '\003\362\000\035'
};
typedef enum MicrosoftWordMPXF MicrosoftWordMPXF;

enum MicrosoftWordMPTs {
	MicrosoftWordMPTsTextEffectShapeUnset = '\000\227\377\376',
	MicrosoftWordMPTsPlainText = '\000\230\000\001',
	MicrosoftWordMPTsStop = '\000\230\000\002',
	MicrosoftWordMPTsTriangleUp = '\000\230\000\003',
	MicrosoftWordMPTsTriangleDown = '\000\230\000\004',
	MicrosoftWordMPTsChevronUp = '\000\230\000\005',
	MicrosoftWordMPTsChevronDown = '\000\230\000\006',
	MicrosoftWordMPTsRingInside = '\000\230\000\007',
	MicrosoftWordMPTsRingOutside = '\000\230\000\010',
	MicrosoftWordMPTsArchUpCurve = '\000\230\000\011',
	MicrosoftWordMPTsArchDownCurve = '\000\230\000\012',
	MicrosoftWordMPTsCircleCurve = '\000\230\000\013',
	MicrosoftWordMPTsButtonCurve = '\000\230\000\014',
	MicrosoftWordMPTsArchUpPour = '\000\230\000\015',
	MicrosoftWordMPTsArchDownPour = '\000\230\000\016',
	MicrosoftWordMPTsCirclePour = '\000\230\000\017',
	MicrosoftWordMPTsButtonPour = '\000\230\000\020',
	MicrosoftWordMPTsCurveUp = '\000\230\000\021',
	MicrosoftWordMPTsCurveDown = '\000\230\000\022',
	MicrosoftWordMPTsCanUp = '\000\230\000\023',
	MicrosoftWordMPTsCanDown = '\000\230\000\024',
	MicrosoftWordMPTsWave1 = '\000\230\000\025',
	MicrosoftWordMPTsWave2 = '\000\230\000\026',
	MicrosoftWordMPTsDoubleWave1 = '\000\230\000\027',
	MicrosoftWordMPTsDoubleWave2 = '\000\230\000\030',
	MicrosoftWordMPTsInflate = '\000\230\000\031',
	MicrosoftWordMPTsDeflate = '\000\230\000\032',
	MicrosoftWordMPTsInflateBottom = '\000\230\000\033',
	MicrosoftWordMPTsDeflateBottom = '\000\230\000\034',
	MicrosoftWordMPTsInflateTop = '\000\230\000\035',
	MicrosoftWordMPTsDeflateTop = '\000\230\000\036',
	MicrosoftWordMPTsDeflateInflate = '\000\230\000\037',
	MicrosoftWordMPTsDeflateInflateDeflate = '\000\230\000 ',
	MicrosoftWordMPTsFadeRight = '\000\230\000!',
	MicrosoftWordMPTsFadeLeft = '\000\230\000\"',
	MicrosoftWordMPTsFadeUp = '\000\230\000#',
	MicrosoftWordMPTsFadeDown = '\000\230\000$',
	MicrosoftWordMPTsSlantUp = '\000\230\000%',
	MicrosoftWordMPTsSlantDown = '\000\230\000&',
	MicrosoftWordMPTsCascadeUp = '\000\230\000\'',
	MicrosoftWordMPTsCascadeDown = '\000\230\000('
};
typedef enum MicrosoftWordMPTs MicrosoftWordMPTs;

enum MicrosoftWordMTxA {
	MicrosoftWordMTxATextEffectAlignmentUnset = '\000\226\377\376',
	MicrosoftWordMTxALeftTextEffectAlignment = '\000\227\000\001',
	MicrosoftWordMTxACenteredTextEffectAlignment = '\000\227\000\002',
	MicrosoftWordMTxARightTextEffectAlignment = '\000\227\000\003',
	MicrosoftWordMTxAJustifyTextEffectAlignment = '\000\227\000\004',
	MicrosoftWordMTxAWordJustifyTextEffectAlignment = '\000\227\000\005',
	MicrosoftWordMTxAStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum MicrosoftWordMTxA MicrosoftWordMTxA;

enum MicrosoftWordMPLd {
	MicrosoftWordMPLdPresetLightingDirectionUnset = '\000\233\377\376',
	MicrosoftWordMPLdLightFromTopLeft = '\000\234\000\001',
	MicrosoftWordMPLdLightFromTop = '\000\234\000\002',
	MicrosoftWordMPLdLightFromTopRight = '\000\234\000\003',
	MicrosoftWordMPLdLightFromLeft = '\000\234\000\004',
	MicrosoftWordMPLdLightFromNone = '\000\234\000\005',
	MicrosoftWordMPLdLightFromRight = '\000\234\000\006',
	MicrosoftWordMPLdLightFromBottomLeft = '\000\234\000\007',
	MicrosoftWordMPLdLightFromBottom = '\000\234\000\010',
	MicrosoftWordMPLdLightFromBottomRight = '\000\234\000\011'
};
typedef enum MicrosoftWordMPLd MicrosoftWordMPLd;

enum MicrosoftWordMlSf {
	MicrosoftWordMlSfLightingSoftnessUnset = '\000\234\377\376',
	MicrosoftWordMlSfLightingDim = '\000\235\000\001',
	MicrosoftWordMlSfLightingNormal = '\000\235\000\002',
	MicrosoftWordMlSfLightingBright = '\000\235\000\003'
};
typedef enum MicrosoftWordMlSf MicrosoftWordMlSf;

enum MicrosoftWordMPMt {
	MicrosoftWordMPMtPresetMaterialUnset = '\000\235\377\376',
	MicrosoftWordMPMtMatte = '\000\236\000\001',
	MicrosoftWordMPMtPlastic = '\000\236\000\002',
	MicrosoftWordMPMtMetal = '\000\236\000\003',
	MicrosoftWordMPMtWireframe = '\000\236\000\004',
	MicrosoftWordMPMtMatte2 = '\000\236\000\005',
	MicrosoftWordMPMtPlastic2 = '\000\236\000\006',
	MicrosoftWordMPMtMetal2 = '\000\236\000\007',
	MicrosoftWordMPMtWarmMatte = '\000\236\000\010',
	MicrosoftWordMPMtTranslucentPowder = '\000\236\000\011',
	MicrosoftWordMPMtPowder = '\000\236\000\012',
	MicrosoftWordMPMtDarkEdge = '\000\236\000\013',
	MicrosoftWordMPMtSoftEdge = '\000\236\000\014',
	MicrosoftWordMPMtMaterialClear = '\000\236\000\015',
	MicrosoftWordMPMtFlat = '\000\236\000\016',
	MicrosoftWordMPMtSoftMetal = '\000\236\000\017'
};
typedef enum MicrosoftWordMPMt MicrosoftWordMPMt;

enum MicrosoftWordMExD {
	MicrosoftWordMExDPresetExtrusionDirectionUnset = '\000\231\377\376',
	MicrosoftWordMExDExtrudeBottomRight = '\000\232\000\001',
	MicrosoftWordMExDExtrudeBottom = '\000\232\000\002',
	MicrosoftWordMExDExtrudeBottomLeft = '\000\232\000\003',
	MicrosoftWordMExDExtrudeRight = '\000\232\000\004',
	MicrosoftWordMExDExtrudeNone = '\000\232\000\005',
	MicrosoftWordMExDExtrudeLeft = '\000\232\000\006',
	MicrosoftWordMExDExtrudeTopRight = '\000\232\000\007',
	MicrosoftWordMExDExtrudeTop = '\000\232\000\010',
	MicrosoftWordMExDExtrudeTopLeft = '\000\232\000\011'
};
typedef enum MicrosoftWordMExD MicrosoftWordMExD;

enum MicrosoftWordM3DF {
	MicrosoftWordM3DFPresetThreeDFormatUnset = '\000\230\377\376',
	MicrosoftWordM3DFFormat1 = '\000\231\000\001',
	MicrosoftWordM3DFFormat2 = '\000\231\000\002',
	MicrosoftWordM3DFFormat3 = '\000\231\000\003',
	MicrosoftWordM3DFFormat4 = '\000\231\000\004',
	MicrosoftWordM3DFFormat5 = '\000\231\000\005',
	MicrosoftWordM3DFFormat6 = '\000\231\000\006',
	MicrosoftWordM3DFFormat7 = '\000\231\000\007',
	MicrosoftWordM3DFFormat8 = '\000\231\000\010',
	MicrosoftWordM3DFFormat9 = '\000\231\000\011',
	MicrosoftWordM3DFFormat10 = '\000\231\000\012',
	MicrosoftWordM3DFFormat11 = '\000\231\000\013',
	MicrosoftWordM3DFFormat12 = '\000\231\000\014',
	MicrosoftWordM3DFFormat13 = '\000\231\000\015',
	MicrosoftWordM3DFFormat14 = '\000\231\000\016',
	MicrosoftWordM3DFFormat15 = '\000\231\000\017',
	MicrosoftWordM3DFFormat16 = '\000\231\000\020',
	MicrosoftWordM3DFFormat17 = '\000\231\000\021',
	MicrosoftWordM3DFFormat18 = '\000\231\000\022',
	MicrosoftWordM3DFFormat19 = '\000\231\000\023',
	MicrosoftWordM3DFFormat20 = '\000\231\000\024'
};
typedef enum MicrosoftWordM3DF MicrosoftWordM3DF;

enum MicrosoftWordMExC {
	MicrosoftWordMExCExtrusionColorUnset = '\000\232\377\376',
	MicrosoftWordMExCExtrusionColorAutomatic = '\000\233\000\001',
	MicrosoftWordMExCExtrusionColorCustom = '\000\233\000\002'
};
typedef enum MicrosoftWordMExC MicrosoftWordMExC;

enum MicrosoftWordMCtT {
	MicrosoftWordMCtTConnectorTypeUnset = '\000h\377\376',
	MicrosoftWordMCtTStraight = '\000i\000\001',
	MicrosoftWordMCtTElbow = '\000i\000\002',
	MicrosoftWordMCtTCurve = '\000i\000\003'
};
typedef enum MicrosoftWordMCtT MicrosoftWordMCtT;

enum MicrosoftWordMHzA {
	MicrosoftWordMHzAHorizontalAnchorUnset = '\000\236\377\376',
	MicrosoftWordMHzAHorizontalAnchorNone = '\000\237\000\001',
	MicrosoftWordMHzAHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum MicrosoftWordMHzA MicrosoftWordMHzA;

enum MicrosoftWordMVtA {
	MicrosoftWordMVtAVerticalAnchorUnset = '\000\237\377\376',
	MicrosoftWordMVtAAnchorTop = '\000\240\000\001',
	MicrosoftWordMVtAAnchorTopBaseline = '\000\240\000\002',
	MicrosoftWordMVtAAnchorMiddle = '\000\240\000\003',
	MicrosoftWordMVtAAnchorBottom = '\000\240\000\004',
	MicrosoftWordMVtAAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum MicrosoftWordMVtA MicrosoftWordMVtA;

enum MicrosoftWordMAsT {
	MicrosoftWordMAsTAutoshapeShapeTypeUnset = '\000i\377\376',
	MicrosoftWordMAsTAutoshapeRectangle = '\000j\000\001',
	MicrosoftWordMAsTAutoshapeParallelogram = '\000j\000\002',
	MicrosoftWordMAsTAutoshapeTrapezoid = '\000j\000\003',
	MicrosoftWordMAsTAutoshapeDiamond = '\000j\000\004',
	MicrosoftWordMAsTAutoshapeRoundedRectangle = '\000j\000\005',
	MicrosoftWordMAsTAutoshapeOctagon = '\000j\000\006',
	MicrosoftWordMAsTAutoshapeIsoscelesTriangle = '\000j\000\007',
	MicrosoftWordMAsTAutoshapeRightTriangle = '\000j\000\010',
	MicrosoftWordMAsTAutoshapeOval = '\000j\000\011',
	MicrosoftWordMAsTAutoshapeHexagon = '\000j\000\012',
	MicrosoftWordMAsTAutoshapeCross = '\000j\000\013',
	MicrosoftWordMAsTAutoshapeRegularPentagon = '\000j\000\014',
	MicrosoftWordMAsTAutoshapeCan = '\000j\000\015',
	MicrosoftWordMAsTAutoshapeCube = '\000j\000\016',
	MicrosoftWordMAsTAutoshapeBevel = '\000j\000\017',
	MicrosoftWordMAsTAutoshapeFoldedCorner = '\000j\000\020',
	MicrosoftWordMAsTAutoshapeSmileyFace = '\000j\000\021',
	MicrosoftWordMAsTAutoshapeDonut = '\000j\000\022',
	MicrosoftWordMAsTAutoshapeNoSymbol = '\000j\000\023',
	MicrosoftWordMAsTAutoshapeBlockArc = '\000j\000\024',
	MicrosoftWordMAsTAutoshapeHeart = '\000j\000\025',
	MicrosoftWordMAsTAutoshapeLightningBolt = '\000j\000\026',
	MicrosoftWordMAsTAutoshapeSun = '\000j\000\027',
	MicrosoftWordMAsTAutoshapeMoon = '\000j\000\030',
	MicrosoftWordMAsTAutoshapeArc = '\000j\000\031',
	MicrosoftWordMAsTAutoshapeDoubleBracket = '\000j\000\032',
	MicrosoftWordMAsTAutoshapeDoubleBrace = '\000j\000\033',
	MicrosoftWordMAsTAutoshapePlaque = '\000j\000\034',
	MicrosoftWordMAsTAutoshapeLeftBracket = '\000j\000\035',
	MicrosoftWordMAsTAutoshapeRightBracket = '\000j\000\036',
	MicrosoftWordMAsTAutoshapeLeftBrace = '\000j\000\037',
	MicrosoftWordMAsTAutoshapeRightBrace = '\000j\000 ',
	MicrosoftWordMAsTAutoshapeRightArrow = '\000j\000!',
	MicrosoftWordMAsTAutoshapeLeftArrow = '\000j\000\"',
	MicrosoftWordMAsTAutoshapeUpArrow = '\000j\000#',
	MicrosoftWordMAsTAutoshapeDownArrow = '\000j\000$',
	MicrosoftWordMAsTAutoshapeLeftRightArrow = '\000j\000%',
	MicrosoftWordMAsTAutoshapeUpDownArrow = '\000j\000&',
	MicrosoftWordMAsTAutoshapeQuadArrow = '\000j\000\'',
	MicrosoftWordMAsTAutoshapeLeftRightUpArrow = '\000j\000(',
	MicrosoftWordMAsTAutoshapeBentArrow = '\000j\000)',
	MicrosoftWordMAsTAutoshapeUTurnArrow = '\000j\000*',
	MicrosoftWordMAsTAutoshapeLeftUpArrow = '\000j\000+',
	MicrosoftWordMAsTAutoshapeBentUpArrow = '\000j\000,',
	MicrosoftWordMAsTAutoshapeCurvedRightArrow = '\000j\000-',
	MicrosoftWordMAsTAutoshapeCurvedLeftArrow = '\000j\000.',
	MicrosoftWordMAsTAutoshapeCurvedUpArrow = '\000j\000/',
	MicrosoftWordMAsTAutoshapeCurvedDownArrow = '\000j\0000',
	MicrosoftWordMAsTAutoshapeStripedRightArrow = '\000j\0001',
	MicrosoftWordMAsTAutoshapeNotchedRightArrow = '\000j\0002',
	MicrosoftWordMAsTAutoshapePentagon = '\000j\0003',
	MicrosoftWordMAsTAutoshapeChevron = '\000j\0004',
	MicrosoftWordMAsTAutoshapeRightArrowCallout = '\000j\0005',
	MicrosoftWordMAsTAutoshapeLeftArrowCallout = '\000j\0006',
	MicrosoftWordMAsTAutoshapeUpArrowCallout = '\000j\0007',
	MicrosoftWordMAsTAutoshapeDownArrowCallout = '\000j\0008',
	MicrosoftWordMAsTAutoshapeLeftRightArrowCallout = '\000j\0009',
	MicrosoftWordMAsTAutoshapeUpDownArrowCallout = '\000j\000:',
	MicrosoftWordMAsTAutoshapeQuadArrowCallout = '\000j\000;',
	MicrosoftWordMAsTAutoshapeCircularArrow = '\000j\000<',
	MicrosoftWordMAsTAutoshapeFlowchartProcess = '\000j\000=',
	MicrosoftWordMAsTAutoshapeFlowchartAlternateProcess = '\000j\000>',
	MicrosoftWordMAsTAutoshapeFlowchartDecision = '\000j\000\?',
	MicrosoftWordMAsTAutoshapeFlowchartData = '\000j\000@',
	MicrosoftWordMAsTAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	MicrosoftWordMAsTAutoshapeFlowchartInternalStorage = '\000j\000B',
	MicrosoftWordMAsTAutoshapeFlowchartDocument = '\000j\000C',
	MicrosoftWordMAsTAutoshapeFlowchartMultiDocument = '\000j\000D',
	MicrosoftWordMAsTAutoshapeFlowchartTerminator = '\000j\000E',
	MicrosoftWordMAsTAutoshapeFlowchartPreparation = '\000j\000F',
	MicrosoftWordMAsTAutoshapeFlowchartManualInput = '\000j\000G',
	MicrosoftWordMAsTAutoshapeFlowchartManualOperation = '\000j\000H',
	MicrosoftWordMAsTAutoshapeFlowchartConnector = '\000j\000I',
	MicrosoftWordMAsTAutoshapeFlowchartOffpageConnector = '\000j\000J',
	MicrosoftWordMAsTAutoshapeFlowchartCard = '\000j\000K',
	MicrosoftWordMAsTAutoshapeFlowchartPunchedTape = '\000j\000L',
	MicrosoftWordMAsTAutoshapeFlowchartSummingJunction = '\000j\000M',
	MicrosoftWordMAsTAutoshapeFlowchartOr = '\000j\000N',
	MicrosoftWordMAsTAutoshapeFlowchartCollate = '\000j\000O',
	MicrosoftWordMAsTAutoshapeFlowchartSort = '\000j\000P',
	MicrosoftWordMAsTAutoshapeFlowchartExtract = '\000j\000Q',
	MicrosoftWordMAsTAutoshapeFlowchartMerge = '\000j\000R',
	MicrosoftWordMAsTAutoshapeFlowchartStoredData = '\000j\000S',
	MicrosoftWordMAsTAutoshapeFlowchartDelay = '\000j\000T',
	MicrosoftWordMAsTAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	MicrosoftWordMAsTAutoshapeFlowchartMagneticDisk = '\000j\000V',
	MicrosoftWordMAsTAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	MicrosoftWordMAsTAutoshapeFlowchartDisplay = '\000j\000X',
	MicrosoftWordMAsTAutoshapeExplosionOne = '\000j\000Y',
	MicrosoftWordMAsTAutoshapeExplosionTwo = '\000j\000Z',
	MicrosoftWordMAsTAutoshapeFourPointStar = '\000j\000[',
	MicrosoftWordMAsTAutoshapeFivePointStar = '\000j\000\\',
	MicrosoftWordMAsTAutoshapeEightPointStar = '\000j\000]',
	MicrosoftWordMAsTAutoshapeSixteenPointStar = '\000j\000^',
	MicrosoftWordMAsTAutoshapeTwentyFourPointStar = '\000j\000_',
	MicrosoftWordMAsTAutoshapeThirtyTwoPointStar = '\000j\000`',
	MicrosoftWordMAsTAutoshapeUpRibbon = '\000j\000a',
	MicrosoftWordMAsTAutoshapeDownRibbon = '\000j\000b',
	MicrosoftWordMAsTAutoshapeCurvedUpRibbon = '\000j\000c',
	MicrosoftWordMAsTAutoshapeCurvedDownRibbon = '\000j\000d',
	MicrosoftWordMAsTAutoshapeVerticalScroll = '\000j\000e',
	MicrosoftWordMAsTAutoshapeHorizontalScroll = '\000j\000f',
	MicrosoftWordMAsTAutoshapeWave = '\000j\000g',
	MicrosoftWordMAsTAutoshapeDoubleWave = '\000j\000h',
	MicrosoftWordMAsTAutoshapeRectangularCallout = '\000j\000i',
	MicrosoftWordMAsTAutoshapeRoundedRectangularCallout = '\000j\000j',
	MicrosoftWordMAsTAutoshapeOvalCallout = '\000j\000k',
	MicrosoftWordMAsTAutoshapeCloudCallout = '\000j\000l',
	MicrosoftWordMAsTAutoshapeLineCalloutOne = '\000j\000m',
	MicrosoftWordMAsTAutoshapeLineCalloutTwo = '\000j\000n',
	MicrosoftWordMAsTAutoshapeLineCalloutThree = '\000j\000o',
	MicrosoftWordMAsTAutoshapeLineCalloutFour = '\000j\000p',
	MicrosoftWordMAsTAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	MicrosoftWordMAsTAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	MicrosoftWordMAsTAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	MicrosoftWordMAsTAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	MicrosoftWordMAsTAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	MicrosoftWordMAsTAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	MicrosoftWordMAsTAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	MicrosoftWordMAsTAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	MicrosoftWordMAsTAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	MicrosoftWordMAsTAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	MicrosoftWordMAsTAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	MicrosoftWordMAsTAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	MicrosoftWordMAsTAutoshapeActionButtonCustom = '\000j\000}',
	MicrosoftWordMAsTAutoshapeActionButtonHome = '\000j\000~',
	MicrosoftWordMAsTAutoshapeActionButtonHelp = '\000j\000\177',
	MicrosoftWordMAsTAutoshapeActionButtonInformation = '\000j\000\200',
	MicrosoftWordMAsTAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	MicrosoftWordMAsTAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	MicrosoftWordMAsTAutoshapeActionButtonBeginning = '\000j\000\203',
	MicrosoftWordMAsTAutoshapeActionButtonEnd = '\000j\000\204',
	MicrosoftWordMAsTAutoshapeActionButtonReturn = '\000j\000\205',
	MicrosoftWordMAsTAutoshapeActionButtonDocument = '\000j\000\206',
	MicrosoftWordMAsTAutoshapeActionButtonSound = '\000j\000\207',
	MicrosoftWordMAsTAutoshapeActionButtonMovie = '\000j\000\210',
	MicrosoftWordMAsTAutoshapeBalloon = '\000j\000\211',
	MicrosoftWordMAsTAutoshapeNotPrimitive = '\000j\000\212',
	MicrosoftWordMAsTAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	MicrosoftWordMAsTAutoshapeLeftRightRibbon = '\000j\000\214',
	MicrosoftWordMAsTAutoshapeDiagonalStripe = '\000j\000\215',
	MicrosoftWordMAsTAutoshapePie = '\000j\000\216',
	MicrosoftWordMAsTAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	MicrosoftWordMAsTAutoshapeDecagon = '\000j\000\220',
	MicrosoftWordMAsTAutoshapeHeptagon = '\000j\000\221',
	MicrosoftWordMAsTAutoshapeDodecagon = '\000j\000\222',
	MicrosoftWordMAsTAutoshapeSixPointsStar = '\000j\000\223',
	MicrosoftWordMAsTAutoshapeSevenPointsStar = '\000j\000\224',
	MicrosoftWordMAsTAutoshapeTenPointsStar = '\000j\000\225',
	MicrosoftWordMAsTAutoshapeTwelvePointsStar = '\000j\000\226',
	MicrosoftWordMAsTAutoshapeRoundOneRectangle = '\000j\000\227',
	MicrosoftWordMAsTAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	MicrosoftWordMAsTAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	MicrosoftWordMAsTAutoshapeSnipRoundRectangle = '\000j\000\232',
	MicrosoftWordMAsTAutoshapeSnipOneRectangle = '\000j\000\233',
	MicrosoftWordMAsTAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	MicrosoftWordMAsTAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	MicrosoftWordMAsTAutoshapeFrame = '\000j\000\236',
	MicrosoftWordMAsTAutoshapeHalfFrame = '\000j\000\237',
	MicrosoftWordMAsTAutoshapeTear = '\000j\000\240',
	MicrosoftWordMAsTAutoshapeChord = '\000j\000\241',
	MicrosoftWordMAsTAutoshapeCorner = '\000j\000\242',
	MicrosoftWordMAsTAutoshapeMathPlus = '\000j\000\243',
	MicrosoftWordMAsTAutoshapeMathMinus = '\000j\000\244',
	MicrosoftWordMAsTAutoshapeMathMultiply = '\000j\000\245',
	MicrosoftWordMAsTAutoshapeMathDivide = '\000j\000\246',
	MicrosoftWordMAsTAutoshapeMathEqual = '\000j\000\247',
	MicrosoftWordMAsTAutoshapeMathNotEqual = '\000j\000\250',
	MicrosoftWordMAsTAutoshapeCornerTabs = '\000j\000\251',
	MicrosoftWordMAsTAutoshapeSquareTabs = '\000j\000\252',
	MicrosoftWordMAsTAutoshapePlaqueTabs = '\000j\000\253',
	MicrosoftWordMAsTAutoshapeGearSix = '\000j\000\254',
	MicrosoftWordMAsTAutoshapeGearNine = '\000j\000\255',
	MicrosoftWordMAsTAutoshapeFunnel = '\000j\000\256',
	MicrosoftWordMAsTAutoshapePieWedge = '\000j\000\257',
	MicrosoftWordMAsTAutoshapeLeftCircularArrow = '\000j\000\260',
	MicrosoftWordMAsTAutoshapeLeftRightCircularArrow = '\000j\000\261',
	MicrosoftWordMAsTAutoshapeSwooshArrow = '\000j\000\262',
	MicrosoftWordMAsTAutoshapeCloud = '\000j\000\263',
	MicrosoftWordMAsTAutoshapeChartX = '\000j\000\264',
	MicrosoftWordMAsTAutoshapeChartStar = '\000j\000\265',
	MicrosoftWordMAsTAutoshapeChartPlus = '\000j\000\266',
	MicrosoftWordMAsTAutoshapeLineInverse = '\000j\000\267'
};
typedef enum MicrosoftWordMAsT MicrosoftWordMAsT;

enum MicrosoftWordMShp {
	MicrosoftWordMShpShapeTypeUnset = '\000\213\377\376',
	MicrosoftWordMShpShapeTypeAuto = '\000\214\000\001',
	MicrosoftWordMShpShapeTypeCallout = '\000\214\000\002',
	MicrosoftWordMShpShapeTypeChart = '\000\214\000\003',
	MicrosoftWordMShpShapeTypeComment = '\000\214\000\004',
	MicrosoftWordMShpShapeTypeFreeForm = '\000\214\000\005',
	MicrosoftWordMShpShapeTypeGroup = '\000\214\000\006',
	MicrosoftWordMShpShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	MicrosoftWordMShpShapeTypeFormControl = '\000\214\000\010',
	MicrosoftWordMShpShapeTypeLine = '\000\214\000\011',
	MicrosoftWordMShpShapeTypeLinkedOLEObject = '\000\214\000\012',
	MicrosoftWordMShpShapeTypeLinkedPicture = '\000\214\000\013',
	MicrosoftWordMShpShapeTypeOLEControl = '\000\214\000\014',
	MicrosoftWordMShpShapeTypePicture = '\000\214\000\015',
	MicrosoftWordMShpShapeTypePlaceHolder = '\000\214\000\016',
	MicrosoftWordMShpShapeTypeWordArt = '\000\214\000\017',
	MicrosoftWordMShpShapeTypeMedia = '\000\214\000\020',
	MicrosoftWordMShpShapeTypeTextBox = '\000\214\000\021',
	MicrosoftWordMShpShapeTypeTable = '\000\214\000\022',
	MicrosoftWordMShpShapeTypeCanvas = '\000\214\000\023',
	MicrosoftWordMShpShapeTypeDiagram = '\000\214\000\024',
	MicrosoftWordMShpShapeTypeInk = '\000\214\000\025',
	MicrosoftWordMShpShapeTypeInkComment = '\000\214\000\026',
	MicrosoftWordMShpShapeTypeSmartartGraphic = '\000\214\000\027',
	MicrosoftWordMShpShapeTypeSlicer = '\000\214\000\030'
};
typedef enum MicrosoftWordMShp MicrosoftWordMShp;

enum MicrosoftWordMCrT {
	MicrosoftWordMCrTColorTypeUnset = '\000j\377\376',
	MicrosoftWordMCrTRGB = '\000k\000\001',
	MicrosoftWordMCrTScheme = '\000k\000\002'
};
typedef enum MicrosoftWordMCrT MicrosoftWordMCrT;

enum MicrosoftWordMPc {
	MicrosoftWordMPcPictureColorTypeUnset = '\000\265\377\376',
	MicrosoftWordMPcPictureColorAutomatic = '\000\266\000\001',
	MicrosoftWordMPcPictureColorGrayScale = '\000\266\000\002',
	MicrosoftWordMPcPictureColorBlackAndWhite = '\000\266\000\003',
	MicrosoftWordMPcPictureColorWatermark = '\000\266\000\004'
};
typedef enum MicrosoftWordMPc MicrosoftWordMPc;

enum MicrosoftWordMCAt {
	MicrosoftWordMCAtAngleUnset = '\000k\377\376',
	MicrosoftWordMCAtAngleAutomatic = '\000l\000\001',
	MicrosoftWordMCAtAngle30 = '\000l\000\002',
	MicrosoftWordMCAtAngle45 = '\000l\000\003',
	MicrosoftWordMCAtAngle60 = '\000l\000\004',
	MicrosoftWordMCAtAngle90 = '\000l\000\005'
};
typedef enum MicrosoftWordMCAt MicrosoftWordMCAt;

enum MicrosoftWordMCDt {
	MicrosoftWordMCDtDropUnset = '\000l\377\376',
	MicrosoftWordMCDtDropCustom = '\000m\000\001',
	MicrosoftWordMCDtDropTop = '\000m\000\002',
	MicrosoftWordMCDtDropCenter = '\000m\000\003',
	MicrosoftWordMCDtDropBottom = '\000m\000\004'
};
typedef enum MicrosoftWordMCDt MicrosoftWordMCDt;

enum MicrosoftWordMCot {
	MicrosoftWordMCotCalloutUnset = '\000m\377\376',
	MicrosoftWordMCotCalloutOne = '\000n\000\001',
	MicrosoftWordMCotCalloutTwo = '\000n\000\002',
	MicrosoftWordMCotCalloutThree = '\000n\000\003',
	MicrosoftWordMCotCalloutFour = '\000n\000\004'
};
typedef enum MicrosoftWordMCot MicrosoftWordMCot;

enum MicrosoftWordTxOr {
	MicrosoftWordTxOrTextOrientationUnset = '\000\215\377\376',
	MicrosoftWordTxOrHorizontal = '\000\216\000\001',
	MicrosoftWordTxOrUpward = '\000\216\000\002',
	MicrosoftWordTxOrDownward = '\000\216\000\003',
	MicrosoftWordTxOrVerticalEastAsian = '\000\216\000\004',
	MicrosoftWordTxOrVertical = '\000\216\000\005',
	MicrosoftWordTxOrHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum MicrosoftWordTxOr MicrosoftWordTxOr;

enum MicrosoftWordMSFr {
	MicrosoftWordMSFrScaleFromTopLeft = '\000o\000\000',
	MicrosoftWordMSFrScaleFromMiddle = '\000o\000\001',
	MicrosoftWordMSFrScaleFromBottomRight = '\000o\000\002'
};
typedef enum MicrosoftWordMSFr MicrosoftWordMSFr;

enum MicrosoftWordMPzC {
	MicrosoftWordMPzCPresetCameraUnset = '\000\256\377\376',
	MicrosoftWordMPzCCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	MicrosoftWordMPzCCameraLegacyObliqueFromTop = '\000\257\000\002',
	MicrosoftWordMPzCCameraLegacyObliqueFromTopright = '\000\257\000\003',
	MicrosoftWordMPzCCameraLegacyObliqueFromLeft = '\000\257\000\004',
	MicrosoftWordMPzCCameraLegacyObliqueFromFront = '\000\257\000\005',
	MicrosoftWordMPzCCameraLegacyObliqueFromRight = '\000\257\000\006',
	MicrosoftWordMPzCCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	MicrosoftWordMPzCCameraLegacyObliqueFromBottom = '\000\257\000\010',
	MicrosoftWordMPzCCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	MicrosoftWordMPzCCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	MicrosoftWordMPzCCameraOrthographic = '\000\257\000\023',
	MicrosoftWordMPzCCameraIsometricFromTopUp = '\000\257\000\024',
	MicrosoftWordMPzCCameraIsometricFromTopDown = '\000\257\000\025',
	MicrosoftWordMPzCCameraIsometricFromBottomUp = '\000\257\000\026',
	MicrosoftWordMPzCCameraIsometricFromBottomDown = '\000\257\000\027',
	MicrosoftWordMPzCCameraIsometricFromLeftUp = '\000\257\000\030',
	MicrosoftWordMPzCCameraIsometricFromLeftDown = '\000\257\000\031',
	MicrosoftWordMPzCCameraIsometricFromRightUp = '\000\257\000\032',
	MicrosoftWordMPzCCameraIsometricFromRightDown = '\000\257\000\033',
	MicrosoftWordMPzCCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	MicrosoftWordMPzCCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	MicrosoftWordMPzCCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	MicrosoftWordMPzCCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	MicrosoftWordMPzCCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	MicrosoftWordMPzCCameraIsometricOffAxis2FromTop = '\000\257\000!',
	MicrosoftWordMPzCCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	MicrosoftWordMPzCCameraIsometricOffAxis3FromRight = '\000\257\000#',
	MicrosoftWordMPzCCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	MicrosoftWordMPzCCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	MicrosoftWordMPzCCameraIsometricOffAxis4FromRight = '\000\257\000&',
	MicrosoftWordMPzCCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	MicrosoftWordMPzCCameraObliqueFromTopLeft = '\000\257\000(',
	MicrosoftWordMPzCCameraObliqueFromTop = '\000\257\000)',
	MicrosoftWordMPzCCameraObliqueFromTopRight = '\000\257\000*',
	MicrosoftWordMPzCCameraObliqueFromLeft = '\000\257\000+',
	MicrosoftWordMPzCCameraObliqueFromRight = '\000\257\000,',
	MicrosoftWordMPzCCameraObliqueFromBottomLeft = '\000\257\000-',
	MicrosoftWordMPzCCameraObliqueFromBottom = '\000\257\000.',
	MicrosoftWordMPzCCameraObliqueFromBottomRight = '\000\257\000/',
	MicrosoftWordMPzCCameraPerspectiveFromFront = '\000\257\0000',
	MicrosoftWordMPzCCameraPerspectiveFromLeft = '\000\257\0001',
	MicrosoftWordMPzCCameraPerspectiveFromRight = '\000\257\0002',
	MicrosoftWordMPzCCameraPerspectiveFromAbove = '\000\257\0003',
	MicrosoftWordMPzCCameraPerspectiveFromBelow = '\000\257\0004',
	MicrosoftWordMPzCCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	MicrosoftWordMPzCCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	MicrosoftWordMPzCCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	MicrosoftWordMPzCCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	MicrosoftWordMPzCCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	MicrosoftWordMPzCCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	MicrosoftWordMPzCCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	MicrosoftWordMPzCCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	MicrosoftWordMPzCCameraPerspectiveRelaxed = '\000\257\000=',
	MicrosoftWordMPzCCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum MicrosoftWordMPzC MicrosoftWordMPzC;

enum MicrosoftWordMLtT {
	MicrosoftWordMLtTLightRigUnset = '\000\257\377\376',
	MicrosoftWordMLtTLightRigFlat1 = '\000\260\000\001',
	MicrosoftWordMLtTLightRigFlat2 = '\000\260\000\002',
	MicrosoftWordMLtTLightRigFlat3 = '\000\260\000\003',
	MicrosoftWordMLtTLightRigFlat4 = '\000\260\000\004',
	MicrosoftWordMLtTLightRigNormal1 = '\000\260\000\005',
	MicrosoftWordMLtTLightRigNormal2 = '\000\260\000\006',
	MicrosoftWordMLtTLightRigNormal3 = '\000\260\000\007',
	MicrosoftWordMLtTLightRigNormal4 = '\000\260\000\010',
	MicrosoftWordMLtTLightRigHarsh1 = '\000\260\000\011',
	MicrosoftWordMLtTLightRigHarsh2 = '\000\260\000\012',
	MicrosoftWordMLtTLightRigHarsh3 = '\000\260\000\013',
	MicrosoftWordMLtTLightRigHarsh4 = '\000\260\000\014',
	MicrosoftWordMLtTLightRigThreePoint = '\000\260\000\015',
	MicrosoftWordMLtTLightRigBalanced = '\000\260\000\016',
	MicrosoftWordMLtTLightRigSoft = '\000\260\000\017',
	MicrosoftWordMLtTLightRigHarsh = '\000\260\000\020',
	MicrosoftWordMLtTLightRigFlood = '\000\260\000\021',
	MicrosoftWordMLtTLightRigContrasting = '\000\260\000\022',
	MicrosoftWordMLtTLightRigMorning = '\000\260\000\023',
	MicrosoftWordMLtTLightRigSunrise = '\000\260\000\024',
	MicrosoftWordMLtTLightRigSunset = '\000\260\000\025',
	MicrosoftWordMLtTLightRigChilly = '\000\260\000\026',
	MicrosoftWordMLtTLightRigFreezing = '\000\260\000\027',
	MicrosoftWordMLtTLightRigFlat = '\000\260\000\030',
	MicrosoftWordMLtTLightRigTwoPoint = '\000\260\000\031',
	MicrosoftWordMLtTLightRigGlow = '\000\260\000\032',
	MicrosoftWordMLtTLightRigBrightRoom = '\000\260\000\033'
};
typedef enum MicrosoftWordMLtT MicrosoftWordMLtT;

enum MicrosoftWordMBlT {
	MicrosoftWordMBlTBevelTypeUnset = '\000\260\377\376',
	MicrosoftWordMBlTBevelNone = '\000\261\000\001',
	MicrosoftWordMBlTBevelRelaxedInset = '\000\261\000\002',
	MicrosoftWordMBlTBevelCircle = '\000\261\000\003',
	MicrosoftWordMBlTBevelSlope = '\000\261\000\004',
	MicrosoftWordMBlTBevelCross = '\000\261\000\005',
	MicrosoftWordMBlTBevelAngle = '\000\261\000\006',
	MicrosoftWordMBlTBevelSoftRound = '\000\261\000\007',
	MicrosoftWordMBlTBevelConvex = '\000\261\000\010',
	MicrosoftWordMBlTBevelCoolSlant = '\000\261\000\011',
	MicrosoftWordMBlTBevelDivot = '\000\261\000\012',
	MicrosoftWordMBlTBevelRiblet = '\000\261\000\013',
	MicrosoftWordMBlTBevelHardEdge = '\000\261\000\014',
	MicrosoftWordMBlTBevelArtDeco = '\000\261\000\015'
};
typedef enum MicrosoftWordMBlT MicrosoftWordMBlT;

enum MicrosoftWordMSSt {
	MicrosoftWordMSStShadowStyleUnset = '\000\261\377\376',
	MicrosoftWordMSStShadowStyleInner = '\000\262\000\001',
	MicrosoftWordMSStShadowStyleOuter = '\000\262\000\002'
};
typedef enum MicrosoftWordMSSt MicrosoftWordMSSt;

enum MicrosoftWordPpgA {
	MicrosoftWordPpgAParagraphAlignmentUnset = '\000\346\377\376',
	MicrosoftWordPpgAParagraphAlignLeft = '\000\347\000\000',
	MicrosoftWordPpgAParagraphAlignCenter = '\000\347\000\001',
	MicrosoftWordPpgAParagraphAlignRight = '\000\347\000\002',
	MicrosoftWordPpgAParagraphAlignJustify = '\000\347\000\003',
	MicrosoftWordPpgAParagraphAlignDistribute = '\000\347\000\004',
	MicrosoftWordPpgAParagraphAlignThai = '\000\347\000\005',
	MicrosoftWordPpgAParagraphAlignJustifyLow = '\000\347\000\006'
};
typedef enum MicrosoftWordPpgA MicrosoftWordPpgA;

enum MicrosoftWordMTSt {
	MicrosoftWordMTStStrikeUnset = '\000\263\377\376',
	MicrosoftWordMTStNoStrike = '\000\264\000\000',
	MicrosoftWordMTStSingleStrike = '\000\264\000\001',
	MicrosoftWordMTStDoubleStrike = '\000\264\000\002'
};
typedef enum MicrosoftWordMTSt MicrosoftWordMTSt;

enum MicrosoftWordMTCa {
	MicrosoftWordMTCaCapsUnset = '\000\264\377\376',
	MicrosoftWordMTCaNoCaps = '\000\265\000\000',
	MicrosoftWordMTCaSmallCaps = '\000\265\000\001',
	MicrosoftWordMTCaAllCaps = '\000\265\000\002'
};
typedef enum MicrosoftWordMTCa MicrosoftWordMTCa;

enum MicrosoftWordMTUl {
	MicrosoftWordMTUlUnderlineUnset = '\003\356\377\376',
	MicrosoftWordMTUlNoUnderline = '\003\357\000\000',
	MicrosoftWordMTUlUnderlineWordsOnly = '\003\357\000\001',
	MicrosoftWordMTUlUnderlineSingleLine = '\003\357\000\002',
	MicrosoftWordMTUlUnderlineDoubleLine = '\003\357\000\003',
	MicrosoftWordMTUlUnderlineHeavyLine = '\003\357\000\004',
	MicrosoftWordMTUlUnderlineDottedLine = '\003\357\000\005',
	MicrosoftWordMTUlUnderlineHeavyDottedLine = '\003\357\000\006',
	MicrosoftWordMTUlUnderlineDashLine = '\003\357\000\007',
	MicrosoftWordMTUlUnderlineHeavyDashLine = '\003\357\000\010',
	MicrosoftWordMTUlUnderlineLongDashLine = '\003\357\000\011',
	MicrosoftWordMTUlUnderlineHeavyLongDashLine = '\003\357\000\012',
	MicrosoftWordMTUlUnderlineDotDashLine = '\003\357\000\013',
	MicrosoftWordMTUlUnderlineHeavyDotDashLine = '\003\357\000\014',
	MicrosoftWordMTUlUnderlineDotDotDashLine = '\003\357\000\015',
	MicrosoftWordMTUlUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	MicrosoftWordMTUlUnderlineWavyLine = '\003\357\000\017',
	MicrosoftWordMTUlUnderlineHeavyWavyLine = '\003\357\000\020',
	MicrosoftWordMTUlUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum MicrosoftWordMTUl MicrosoftWordMTUl;

enum MicrosoftWordMTTA {
	MicrosoftWordMTTATabUnset = '\000\266\377\376',
	MicrosoftWordMTTALeftTab = '\000\267\000\000',
	MicrosoftWordMTTACenterTab = '\000\267\000\001',
	MicrosoftWordMTTARightTab = '\000\267\000\002',
	MicrosoftWordMTTADecimalTab = '\000\267\000\003'
};
typedef enum MicrosoftWordMTTA MicrosoftWordMTTA;

enum MicrosoftWordMTCW {
	MicrosoftWordMTCWCharacterWrapUnset = '\000\267\377\376',
	MicrosoftWordMTCWNoCharacterWrap = '\000\270\000\000',
	MicrosoftWordMTCWStandardCharacterWrap = '\000\270\000\001',
	MicrosoftWordMTCWStrictCharacterWrap = '\000\270\000\002',
	MicrosoftWordMTCWCustomCharacterWrap = '\000\270\000\003'
};
typedef enum MicrosoftWordMTCW MicrosoftWordMTCW;

enum MicrosoftWordMTFA {
	MicrosoftWordMTFAFontAlignmentUnset = '\000\270\377\376',
	MicrosoftWordMTFAAutomaticAlignment = '\000\271\000\000',
	MicrosoftWordMTFATopAlignment = '\000\271\000\001',
	MicrosoftWordMTFACenterAlignment = '\000\271\000\002',
	MicrosoftWordMTFABaselineAlignment = '\000\271\000\003',
	MicrosoftWordMTFABottomAlignment = '\000\271\000\004'
};
typedef enum MicrosoftWordMTFA MicrosoftWordMTFA;

enum MicrosoftWordPAtS {
	MicrosoftWordPAtSAutoSizeUnset = '\000\344\377\376',
	MicrosoftWordPAtSAutoSizeNone = '\000\345\000\000',
	MicrosoftWordPAtSShapeToFitText = '\000\345\000\001',
	MicrosoftWordPAtSTextToFitShape = '\000\345\000\002'
};
typedef enum MicrosoftWordPAtS MicrosoftWordPAtS;

enum MicrosoftWordMPFo {
	MicrosoftWordMPFoPathTypeUnset = '\000\272\377\376',
	MicrosoftWordMPFoNoPathType = '\000\273\000\000',
	MicrosoftWordMPFoPathType1 = '\000\273\000\001',
	MicrosoftWordMPFoPathType2 = '\000\273\000\002',
	MicrosoftWordMPFoPathType3 = '\000\273\000\003',
	MicrosoftWordMPFoPathType4 = '\000\273\000\004'
};
typedef enum MicrosoftWordMPFo MicrosoftWordMPFo;

enum MicrosoftWordMWFo {
	MicrosoftWordMWFoWarpFormatUnset = '\000\273\377\376',
	MicrosoftWordMWFoWarpFormat1 = '\000\274\000\000',
	MicrosoftWordMWFoWarpFormat2 = '\000\274\000\001',
	MicrosoftWordMWFoWarpFormat3 = '\000\274\000\002',
	MicrosoftWordMWFoWarpFormat4 = '\000\274\000\003',
	MicrosoftWordMWFoWarpFormat5 = '\000\274\000\004',
	MicrosoftWordMWFoWarpFormat6 = '\000\274\000\005',
	MicrosoftWordMWFoWarpFormat7 = '\000\274\000\006',
	MicrosoftWordMWFoWarpFormat8 = '\000\274\000\007',
	MicrosoftWordMWFoWarpFormat9 = '\000\274\000\010',
	MicrosoftWordMWFoWarpFormat10 = '\000\274\000\011',
	MicrosoftWordMWFoWarpFormat11 = '\000\274\000\012',
	MicrosoftWordMWFoWarpFormat12 = '\000\274\000\013',
	MicrosoftWordMWFoWarpFormat13 = '\000\274\000\014',
	MicrosoftWordMWFoWarpFormat14 = '\000\274\000\015',
	MicrosoftWordMWFoWarpFormat15 = '\000\274\000\016',
	MicrosoftWordMWFoWarpFormat16 = '\000\274\000\017',
	MicrosoftWordMWFoWarpFormat17 = '\000\274\000\020',
	MicrosoftWordMWFoWarpFormat18 = '\000\274\000\021',
	MicrosoftWordMWFoWarpFormat19 = '\000\274\000\022',
	MicrosoftWordMWFoWarpFormat20 = '\000\274\000\023',
	MicrosoftWordMWFoWarpFormat21 = '\000\274\000\024',
	MicrosoftWordMWFoWarpFormat22 = '\000\274\000\025',
	MicrosoftWordMWFoWarpFormat23 = '\000\274\000\026',
	MicrosoftWordMWFoWarpFormat24 = '\000\274\000\027',
	MicrosoftWordMWFoWarpFormat25 = '\000\274\000\030',
	MicrosoftWordMWFoWarpFormat26 = '\000\274\000\031',
	MicrosoftWordMWFoWarpFormat27 = '\000\274\000\032',
	MicrosoftWordMWFoWarpFormat28 = '\000\274\000\033',
	MicrosoftWordMWFoWarpFormat29 = '\000\274\000\034',
	MicrosoftWordMWFoWarpFormat30 = '\000\274\000\035',
	MicrosoftWordMWFoWarpFormat31 = '\000\274\000\036',
	MicrosoftWordMWFoWarpFormat32 = '\000\274\000\037',
	MicrosoftWordMWFoWarpFormat33 = '\000\274\000 ',
	MicrosoftWordMWFoWarpFormat34 = '\000\274\000!',
	MicrosoftWordMWFoWarpFormat35 = '\000\274\000\"',
	MicrosoftWordMWFoWarpFormat36 = '\000\274\000#'
};
typedef enum MicrosoftWordMWFo MicrosoftWordMWFo;

enum MicrosoftWordPCgC {
	MicrosoftWordPCgCCaseSentence = '\000\344\000\001',
	MicrosoftWordPCgCCaseLower = '\000\344\000\002',
	MicrosoftWordPCgCCaseUpper = '\000\344\000\003',
	MicrosoftWordPCgCCaseTitle = '\000\344\000\004',
	MicrosoftWordPCgCCaseToggle = '\000\344\000\005'
};
typedef enum MicrosoftWordPCgC MicrosoftWordPCgC;

enum MicrosoftWordMDTF {
	MicrosoftWordMDTFDateTimeFormatUnset = '\000\275\377\376',
	MicrosoftWordMDTFDateTimeFormatMdyy = '\000\276\000\001',
	MicrosoftWordMDTFDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	MicrosoftWordMDTFDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	MicrosoftWordMDTFDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	MicrosoftWordMDTFDateTimeFormatDMMMyy = '\000\276\000\005',
	MicrosoftWordMDTFDateTimeFormatMMMMyy = '\000\276\000\006',
	MicrosoftWordMDTFDateTimeFormatMMyy = '\000\276\000\007',
	MicrosoftWordMDTFDateTimeFormatMMddyyHmm = '\000\276\000\010',
	MicrosoftWordMDTFDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	MicrosoftWordMDTFDateTimeFormatHmm = '\000\276\000\012',
	MicrosoftWordMDTFDateTimeFormatHmmss = '\000\276\000\013',
	MicrosoftWordMDTFDateTimeFormatHmmAMPM = '\000\276\000\014',
	MicrosoftWordMDTFDateTimeFormatHmmssAMPM = '\000\276\000\015',
	MicrosoftWordMDTFDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum MicrosoftWordMDTF MicrosoftWordMDTF;

enum MicrosoftWordMSET {
	MicrosoftWordMSETSoftEdgeUnset = '\000\277\377\376',
	MicrosoftWordMSETNoSoftEdge = '\000\300\000\000',
	MicrosoftWordMSETSoftEdgeType1 = '\000\300\000\001',
	MicrosoftWordMSETSoftEdgeType2 = '\000\300\000\002',
	MicrosoftWordMSETSoftEdgeType3 = '\000\300\000\003',
	MicrosoftWordMSETSoftEdgeType4 = '\000\300\000\004',
	MicrosoftWordMSETSoftEdgeType5 = '\000\300\000\005',
	MicrosoftWordMSETSoftEdgeType6 = '\000\300\000\006'
};
typedef enum MicrosoftWordMSET MicrosoftWordMSET;

enum MicrosoftWordMCSI {
	MicrosoftWordMCSIFirstDarkSchemeColor = '\000\301\000\001',
	MicrosoftWordMCSIFirstLightSchemeColor = '\000\301\000\002',
	MicrosoftWordMCSISecondDarkSchemeColor = '\000\301\000\003',
	MicrosoftWordMCSISecondLightSchemeColor = '\000\301\000\004',
	MicrosoftWordMCSIFirstAccentSchemeColor = '\000\301\000\005',
	MicrosoftWordMCSISecondAccentSchemeColor = '\000\301\000\006',
	MicrosoftWordMCSIThirdAccentSchemeColor = '\000\301\000\007',
	MicrosoftWordMCSIFourthAccentSchemeColor = '\000\301\000\010',
	MicrosoftWordMCSIFifthAccentSchemeColor = '\000\301\000\011',
	MicrosoftWordMCSISixthAccentSchemeColor = '\000\301\000\012',
	MicrosoftWordMCSIHyperlinkSchemeColor = '\000\301\000\013',
	MicrosoftWordMCSIFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum MicrosoftWordMCSI MicrosoftWordMCSI;

enum MicrosoftWordMCoI {
	MicrosoftWordMCoIThemeColorUnset = '\000\301\377\376',
	MicrosoftWordMCoINoThemeColor = '\000\302\000\000',
	MicrosoftWordMCoIFirstDarkThemeColor = '\000\302\000\001',
	MicrosoftWordMCoIFirstLightThemeColor = '\000\302\000\002',
	MicrosoftWordMCoISecondDarkThemeColor = '\000\302\000\003',
	MicrosoftWordMCoISecondLightThemeColor = '\000\302\000\004',
	MicrosoftWordMCoIFirstAccentThemeColor = '\000\302\000\005',
	MicrosoftWordMCoISecondAccentThemeColor = '\000\302\000\006',
	MicrosoftWordMCoIThirdAccentThemeColor = '\000\302\000\007',
	MicrosoftWordMCoIFourthAccentThemeColor = '\000\302\000\010',
	MicrosoftWordMCoIFifthAccentThemeColor = '\000\302\000\011',
	MicrosoftWordMCoISixthAccentThemeColor = '\000\302\000\012',
	MicrosoftWordMCoIHyperlinkThemeColor = '\000\302\000\013',
	MicrosoftWordMCoIFollowedHyperlinkThemeColor = '\000\302\000\014',
	MicrosoftWordMCoIFirstTextThemeColor = '\000\302\000\015',
	MicrosoftWordMCoIFirstBackgroundThemeColor = '\000\302\000\016',
	MicrosoftWordMCoISecondTextThemeColor = '\000\302\000\017',
	MicrosoftWordMCoISecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum MicrosoftWordMCoI MicrosoftWordMCoI;

enum MicrosoftWordMFLI {
	MicrosoftWordMFLIThemeFontLatin = '\000\303\000\001',
	MicrosoftWordMFLIThemeFontComplexScript = '\000\303\000\002',
	MicrosoftWordMFLIThemeFontHighAnsi = '\000\303\000\003',
	MicrosoftWordMFLIThemeFontEastAsian = '\000\303\000\004'
};
typedef enum MicrosoftWordMFLI MicrosoftWordMFLI;

enum MicrosoftWordMSSI {
	MicrosoftWordMSSIShapeStyleUnset = '\000\303\377\376',
	MicrosoftWordMSSIShapeNotAPreset = '\000\304\000\000',
	MicrosoftWordMSSIShapePreset1 = '\000\304\000\001',
	MicrosoftWordMSSIShapePreset2 = '\000\304\000\002',
	MicrosoftWordMSSIShapePreset3 = '\000\304\000\003',
	MicrosoftWordMSSIShapePreset4 = '\000\304\000\004',
	MicrosoftWordMSSIShapePreset5 = '\000\304\000\005',
	MicrosoftWordMSSIShapePreset6 = '\000\304\000\006',
	MicrosoftWordMSSIShapePreset7 = '\000\304\000\007',
	MicrosoftWordMSSIShapePreset8 = '\000\304\000\010',
	MicrosoftWordMSSIShapePreset9 = '\000\304\000\011',
	MicrosoftWordMSSIShapePreset10 = '\000\304\000\012',
	MicrosoftWordMSSIShapePreset11 = '\000\304\000\013',
	MicrosoftWordMSSIShapePreset12 = '\000\304\000\014',
	MicrosoftWordMSSIShapePreset13 = '\000\304\000\015',
	MicrosoftWordMSSIShapePreset14 = '\000\304\000\016',
	MicrosoftWordMSSIShapePreset15 = '\000\304\000\017',
	MicrosoftWordMSSIShapePreset16 = '\000\304\000\020',
	MicrosoftWordMSSIShapePreset17 = '\000\304\000\021',
	MicrosoftWordMSSIShapePreset18 = '\000\304\000\022',
	MicrosoftWordMSSIShapePreset19 = '\000\304\000\023',
	MicrosoftWordMSSIShapePreset20 = '\000\304\000\024',
	MicrosoftWordMSSIShapePreset21 = '\000\304\000\025',
	MicrosoftWordMSSIShapePreset22 = '\000\304\000\026',
	MicrosoftWordMSSIShapePreset23 = '\000\304\000\027',
	MicrosoftWordMSSIShapePreset24 = '\000\304\000\030',
	MicrosoftWordMSSIShapePreset25 = '\000\304\000\031',
	MicrosoftWordMSSIShapePreset26 = '\000\304\000\032',
	MicrosoftWordMSSIShapePreset27 = '\000\304\000\033',
	MicrosoftWordMSSIShapePreset28 = '\000\304\000\034',
	MicrosoftWordMSSIShapePreset29 = '\000\304\000\035',
	MicrosoftWordMSSIShapePreset30 = '\000\304\000\036',
	MicrosoftWordMSSIShapePreset31 = '\000\304\000\037',
	MicrosoftWordMSSIShapePreset32 = '\000\304\000 ',
	MicrosoftWordMSSIShapePreset33 = '\000\304\000!',
	MicrosoftWordMSSIShapePreset34 = '\000\304\000\"',
	MicrosoftWordMSSIShapePreset35 = '\000\304\000#',
	MicrosoftWordMSSIShapePreset36 = '\000\304\000$',
	MicrosoftWordMSSIShapePreset37 = '\000\304\000%',
	MicrosoftWordMSSIShapePreset38 = '\000\304\000&',
	MicrosoftWordMSSIShapePreset39 = '\000\304\000\'',
	MicrosoftWordMSSIShapePreset40 = '\000\304\000(',
	MicrosoftWordMSSIShapePreset41 = '\000\304\000)',
	MicrosoftWordMSSIShapePreset42 = '\000\304\000*',
	MicrosoftWordMSSILinePreset1 = '\000\304\'\021',
	MicrosoftWordMSSILinePreset2 = '\000\304\'\022',
	MicrosoftWordMSSILinePreset3 = '\000\304\'\023',
	MicrosoftWordMSSILinePreset4 = '\000\304\'\024',
	MicrosoftWordMSSILinePreset5 = '\000\304\'\025',
	MicrosoftWordMSSILinePreset6 = '\000\304\'\026',
	MicrosoftWordMSSILinePreset7 = '\000\304\'\027',
	MicrosoftWordMSSILinePreset8 = '\000\304\'\030',
	MicrosoftWordMSSILinePreset9 = '\000\304\'\031',
	MicrosoftWordMSSILinePreset10 = '\000\304\'\032',
	MicrosoftWordMSSILinePreset11 = '\000\304\'\033',
	MicrosoftWordMSSILinePreset12 = '\000\304\'\034',
	MicrosoftWordMSSILinePreset13 = '\000\304\'\035',
	MicrosoftWordMSSILinePreset14 = '\000\304\'\036',
	MicrosoftWordMSSILinePreset15 = '\000\304\'\037',
	MicrosoftWordMSSILinePreset16 = '\000\304\' ',
	MicrosoftWordMSSILinePreset17 = '\000\304\'!',
	MicrosoftWordMSSILinePreset18 = '\000\304\'\"',
	MicrosoftWordMSSILinePreset19 = '\000\304\'#',
	MicrosoftWordMSSILinePreset20 = '\000\304\'$',
	MicrosoftWordMSSILinePreset21 = '\000\304\'%'
};
typedef enum MicrosoftWordMSSI MicrosoftWordMSSI;

enum MicrosoftWordMBSI {
	MicrosoftWordMBSIBackgroundUnset = '\000\304\377\376',
	MicrosoftWordMBSIBackgroundNotAPreset = '\000\305\000\000',
	MicrosoftWordMBSIBackgroundPreset1 = '\000\305\000\001',
	MicrosoftWordMBSIBackgroundPreset2 = '\000\305\000\002',
	MicrosoftWordMBSIBackgroundPreset3 = '\000\305\000\003',
	MicrosoftWordMBSIBackgroundPreset4 = '\000\305\000\004',
	MicrosoftWordMBSIBackgroundPreset5 = '\000\305\000\005',
	MicrosoftWordMBSIBackgroundPreset6 = '\000\305\000\006',
	MicrosoftWordMBSIBackgroundPreset7 = '\000\305\000\007',
	MicrosoftWordMBSIBackgroundPreset8 = '\000\305\000\010',
	MicrosoftWordMBSIBackgroundPreset9 = '\000\305\000\011',
	MicrosoftWordMBSIBackgroundPreset10 = '\000\305\000\012',
	MicrosoftWordMBSIBackgroundPreset11 = '\000\305\000\013',
	MicrosoftWordMBSIBackgroundPreset12 = '\000\305\000\014'
};
typedef enum MicrosoftWordMBSI MicrosoftWordMBSI;

enum MicrosoftWordPDrT {
	MicrosoftWordPDrTTextDirectionUnset = '\000\352\377\376',
	MicrosoftWordPDrTLeftToRight = '\000\353\000\001',
	MicrosoftWordPDrTRightToLeft = '\000\353\000\002'
};
typedef enum MicrosoftWordPDrT MicrosoftWordPDrT;

enum MicrosoftWordPBtT {
	MicrosoftWordPBtTBulletTypeUnset = '\000\347\377\376',
	MicrosoftWordPBtTBulletTypeNone = '\000\350\000\000',
	MicrosoftWordPBtTBulletTypeUnnumbered = '\000\350\000\001',
	MicrosoftWordPBtTBulletTypeNumbered = '\000\350\000\002',
	MicrosoftWordPBtTPictureBulletType = '\000\350\000\003'
};
typedef enum MicrosoftWordPBtT MicrosoftWordPBtT;

enum MicrosoftWordPBtS {
	MicrosoftWordPBtSNumberedBulletStyleUnset = '\000\350\377\376',
	MicrosoftWordPBtSNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	MicrosoftWordPBtSNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	MicrosoftWordPBtSNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	MicrosoftWordPBtSNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	MicrosoftWordPBtSNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	MicrosoftWordPBtSNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	MicrosoftWordPBtSNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	MicrosoftWordPBtSNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	MicrosoftWordPBtSNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	MicrosoftWordPBtSNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	MicrosoftWordPBtSNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	MicrosoftWordPBtSNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	MicrosoftWordPBtSNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	MicrosoftWordPBtSNumberedBulletStyleArabicPlain = '\000\351\000\015',
	MicrosoftWordPBtSNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	MicrosoftWordPBtSNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	MicrosoftWordPBtSNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	MicrosoftWordPBtSNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	MicrosoftWordPBtSNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	MicrosoftWordPBtSNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	MicrosoftWordPBtSNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	MicrosoftWordPBtSNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	MicrosoftWordPBtSNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	MicrosoftWordPBtSNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	MicrosoftWordPBtSNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	MicrosoftWordPBtSNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	MicrosoftWordPBtSNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	MicrosoftWordPBtSNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	MicrosoftWordPBtSNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	MicrosoftWordPBtSNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	MicrosoftWordPBtSNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	MicrosoftWordPBtSNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	MicrosoftWordPBtSNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	MicrosoftWordPBtSNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	MicrosoftWordPBtSNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	MicrosoftWordPBtSNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	MicrosoftWordPBtSNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	MicrosoftWordPBtSNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	MicrosoftWordPBtSNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	MicrosoftWordPBtSNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	MicrosoftWordPBtSNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum MicrosoftWordPBtS MicrosoftWordPBtS;

enum MicrosoftWordPTSp {
	MicrosoftWordPTSpTabstopUnset = '\000\364\377\376',
	MicrosoftWordPTSpTabstopLeft = '\000\365\000\001',
	MicrosoftWordPTSpTabstopCenter = '\000\365\000\002',
	MicrosoftWordPTSpTabstopRight = '\000\365\000\003',
	MicrosoftWordPTSpTabstopDecimal = '\000\365\000\004'
};
typedef enum MicrosoftWordPTSp MicrosoftWordPTSp;

enum MicrosoftWordMRfT {
	MicrosoftWordMRfTReflectionUnset = '\003\351\377\376',
	MicrosoftWordMRfTReflectionTypeNone = '\003\352\000\000',
	MicrosoftWordMRfTReflectionType1 = '\003\352\000\001',
	MicrosoftWordMRfTReflectionType2 = '\003\352\000\002',
	MicrosoftWordMRfTReflectionType3 = '\003\352\000\003',
	MicrosoftWordMRfTReflectionType4 = '\003\352\000\004',
	MicrosoftWordMRfTReflectionType5 = '\003\352\000\005',
	MicrosoftWordMRfTReflectionType6 = '\003\352\000\006',
	MicrosoftWordMRfTReflectionType7 = '\003\352\000\007',
	MicrosoftWordMRfTReflectionType8 = '\003\352\000\010',
	MicrosoftWordMRfTReflectionType9 = '\003\352\000\011'
};
typedef enum MicrosoftWordMRfT MicrosoftWordMRfT;

enum MicrosoftWordMTtA {
	MicrosoftWordMTtATextureUnset = '\003\352\377\376',
	MicrosoftWordMTtATextureTopLeft = '\003\353\000\000',
	MicrosoftWordMTtATextureTop = '\003\353\000\001',
	MicrosoftWordMTtATextureTopRight = '\003\353\000\002',
	MicrosoftWordMTtATextureLeft = '\003\353\000\003',
	MicrosoftWordMTtATextureCenter = '\003\353\000\004',
	MicrosoftWordMTtATextureRight = '\003\353\000\005',
	MicrosoftWordMTtATextureBottomLeft = '\003\353\000\006',
	MicrosoftWordMTtATextureBotton = '\003\353\000\007',
	MicrosoftWordMTtATextureBottomRight = '\003\353\000\010'
};
typedef enum MicrosoftWordMTtA MicrosoftWordMTtA;

enum MicrosoftWordPBlA {
	MicrosoftWordPBlATextBaselineAlignmentUnset = '\003\353\377\376',
	MicrosoftWordPBlATextBaselineAlignBaseline = '\003\354\000\001',
	MicrosoftWordPBlATextBaselineAlignTop = '\003\354\000\002',
	MicrosoftWordPBlATextBaselineAlignCenter = '\003\354\000\003',
	MicrosoftWordPBlATextBaselineAlignEastAsian50 = '\003\354\000\004',
	MicrosoftWordPBlATextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum MicrosoftWordPBlA MicrosoftWordPBlA;

enum MicrosoftWordMCbF {
	MicrosoftWordMCbFClipboardFormatUnset = '\003\354\377\376',
	MicrosoftWordMCbFNativeClipboardFormat = '\003\355\000\001',
	MicrosoftWordMCbFHTMlClipboardFormat = '\003\355\000\002',
	MicrosoftWordMCbFRTFClipboardFormat = '\003\355\000\003',
	MicrosoftWordMCbFPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum MicrosoftWordMCbF MicrosoftWordMCbF;

enum MicrosoftWordMTiP {
	MicrosoftWordMTiPInsertBefore = '\003\356\000\000',
	MicrosoftWordMTiPInsertAfter = '\003\356\000\001'
};
typedef enum MicrosoftWordMTiP MicrosoftWordMTiP;

enum MicrosoftWordMPiT {
	MicrosoftWordMPiTSaveAsDefault = '\003\362\377\376',
	MicrosoftWordMPiTSaveAsPNGFile = '\003\363\000\000',
	MicrosoftWordMPiTSaveAsBMPFile = '\003\363\000\001',
	MicrosoftWordMPiTSaveAsGIFFile = '\003\363\000\002',
	MicrosoftWordMPiTSaveAsJPGFile = '\003\363\000\003',
	MicrosoftWordMPiTSaveAsPDFFile = '\003\363\000\004'
};
typedef enum MicrosoftWordMPiT MicrosoftWordMPiT;

enum MicrosoftWordMPeT {
	MicrosoftWordMPeTNoEffect = '\003\364\000\000',
	MicrosoftWordMPeTEffectBackgroundRemoval = '\003\364\000\001',
	MicrosoftWordMPeTEffectBlur = '\003\364\000\002',
	MicrosoftWordMPeTEffectBrightnessContrast = '\003\364\000\003',
	MicrosoftWordMPeTEffectCement = '\003\364\000\004',
	MicrosoftWordMPeTEffectCrisscrossEtching = '\003\364\000\005',
	MicrosoftWordMPeTEffectChalkSketch = '\003\364\000\006',
	MicrosoftWordMPeTEffectColorTemperature = '\003\364\000\007',
	MicrosoftWordMPeTEffectCutout = '\003\364\000\010',
	MicrosoftWordMPeTEffectFilmGrain = '\003\364\000\011',
	MicrosoftWordMPeTEffectGlass = '\003\364\000\012',
	MicrosoftWordMPeTEffectGlowDiffused = '\003\364\000\013',
	MicrosoftWordMPeTEffectGlowEdges = '\003\364\000\014',
	MicrosoftWordMPeTEffectLightScreen = '\003\364\000\015',
	MicrosoftWordMPeTEffectLineDrawing = '\003\364\000\016',
	MicrosoftWordMPeTEffectMarker = '\003\364\000\017',
	MicrosoftWordMPeTEffectMosiaicBubbles = '\003\364\000\020',
	MicrosoftWordMPeTEffectPaintBrush = '\003\364\000\021',
	MicrosoftWordMPeTEffectPaintStrokes = '\003\364\000\022',
	MicrosoftWordMPeTEffectPastelsSmooth = '\003\364\000\023',
	MicrosoftWordMPeTEffectPencilGrayscale = '\003\364\000\024',
	MicrosoftWordMPeTEffectPencilSketch = '\003\364\000\025',
	MicrosoftWordMPeTEffectPhotocopy = '\003\364\000\026',
	MicrosoftWordMPeTEffectPlasticWrap = '\003\364\000\027',
	MicrosoftWordMPeTEffectSaturation = '\003\364\000\030',
	MicrosoftWordMPeTEffectSharpenSoften = '\003\364\000\031',
	MicrosoftWordMPeTEffectTexturizer = '\003\364\000\032',
	MicrosoftWordMPeTEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum MicrosoftWordMPeT MicrosoftWordMPeT;

enum MicrosoftWordMSgT {
	MicrosoftWordMSgTLine = '\000\217\000\000',
	MicrosoftWordMSgTCurve = '\000\217\000\001'
};
typedef enum MicrosoftWordMSgT MicrosoftWordMSgT;

enum MicrosoftWordMEdT {
	MicrosoftWordMEdTAuto = '\000\220\000\000',
	MicrosoftWordMEdTCorner = '\000\220\000\001',
	MicrosoftWordMEdTSmooth = '\000\220\000\002',
	MicrosoftWordMEdTSymmetric = '\000\220\000\003'
};
typedef enum MicrosoftWordMEdT MicrosoftWordMEdT;

enum MicrosoftWordSANP {
	MicrosoftWordSANPDefaultNodePosition = '\003\365\000\001',
	MicrosoftWordSANPAfterNode = '\003\365\000\002',
	MicrosoftWordSANPBeforeNode = '\003\365\000\003',
	MicrosoftWordSANPAboveNode = '\003\365\000\004',
	MicrosoftWordSANPBelowNode = '\003\365\000\005'
};
typedef enum MicrosoftWordSANP MicrosoftWordSANP;

enum MicrosoftWordSANT {
	MicrosoftWordSANTDefaultNode = '\003\366\000\001',
	MicrosoftWordSANTAssistantNode = '\003\366\000\002'
};
typedef enum MicrosoftWordSANT MicrosoftWordSANT;

enum MicrosoftWordMOCT {
	MicrosoftWordMOCTOrgChartLayoutUnset = '\003\366\377\376',
	MicrosoftWordMOCTOrgChartLayoutStandard = '\003\367\000\001',
	MicrosoftWordMOCTOrgChartLayoutBothHanging = '\003\367\000\002',
	MicrosoftWordMOCTOrgChartLayoutLeftHanging = '\003\367\000\003',
	MicrosoftWordMOCTOrgChartLayoutRightHanging = '\003\367\000\004',
	MicrosoftWordMOCTOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum MicrosoftWordMOCT MicrosoftWordMOCT;

enum MicrosoftWordMAlC {
	MicrosoftWordMAlCAlignLefts = '\000\000\000\000',
	MicrosoftWordMAlCAlignCenters = '\000\000\000\001',
	MicrosoftWordMAlCAlignRights = '\000\000\000\002',
	MicrosoftWordMAlCAlignTops = '\000\000\000\003',
	MicrosoftWordMAlCAlignMiddles = '\000\000\000\004',
	MicrosoftWordMAlCAlignBottoms = '\000\000\000\005'
};
typedef enum MicrosoftWordMAlC MicrosoftWordMAlC;

enum MicrosoftWordMDsC {
	MicrosoftWordMDsCDistributeHorizontally = '\000\000\000\000',
	MicrosoftWordMDsCDistributeVertically = '\000\000\000\001'
};
typedef enum MicrosoftWordMDsC MicrosoftWordMDsC;

enum MicrosoftWordMOrT {
	MicrosoftWordMOrTOrientationUnset = '\000\214\377\376',
	MicrosoftWordMOrTHorizontalOrientation = '\000\215\000\001',
	MicrosoftWordMOrTVerticalOrientation = '\000\215\000\002'
};
typedef enum MicrosoftWordMOrT MicrosoftWordMOrT;

enum MicrosoftWordMZoC {
	MicrosoftWordMZoCBringShapeToFront = '\000p\000\000',
	MicrosoftWordMZoCSendShapeToBack = '\000p\000\001',
	MicrosoftWordMZoCBringShapeForward = '\000p\000\002',
	MicrosoftWordMZoCSendShapeBackward = '\000p\000\003',
	MicrosoftWordMZoCBringShapeInFrontOfText = '\000p\000\004',
	MicrosoftWordMZoCSendShapeBehindText = '\000p\000\005'
};
typedef enum MicrosoftWordMZoC MicrosoftWordMZoC;

enum MicrosoftWordMFlC {
	MicrosoftWordMFlCFlipHorizontal = '\000q\000\000',
	MicrosoftWordMFlCFlipVertical = '\000q\000\001'
};
typedef enum MicrosoftWordMFlC MicrosoftWordMFlC;

enum MicrosoftWordMTri {
	MicrosoftWordMTriTrue = '\000\240\377\377',
	MicrosoftWordMTriFalse = '\000\241\000\000',
	MicrosoftWordMTriCTrue = '\000\241\000\001',
	MicrosoftWordMTriToggle = '\000\240\377\375',
	MicrosoftWordMTriTriStateUnset = '\000\240\377\376'
};
typedef enum MicrosoftWordMTri MicrosoftWordMTri;

enum MicrosoftWordMBW {
	MicrosoftWordMBWBlackAndWhiteUnset = '\000\254\377\376',
	MicrosoftWordMBWBlackAndWhiteModeAutomatic = '\000\255\000\001',
	MicrosoftWordMBWBlackAndWhiteModeGrayScale = '\000\255\000\002',
	MicrosoftWordMBWBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	MicrosoftWordMBWBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	MicrosoftWordMBWBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	MicrosoftWordMBWBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	MicrosoftWordMBWBlackAndWhiteModeHighContrast = '\000\255\000\007',
	MicrosoftWordMBWBlackAndWhiteModeBlack = '\000\255\000\010',
	MicrosoftWordMBWBlackAndWhiteModeWhite = '\000\255\000\011',
	MicrosoftWordMBWBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum MicrosoftWordMBW MicrosoftWordMBW;

enum MicrosoftWordMBPS {
	MicrosoftWordMBPSBarLeft = '\000r\000\000',
	MicrosoftWordMBPSBarTop = '\000r\000\001',
	MicrosoftWordMBPSBarRight = '\000r\000\002',
	MicrosoftWordMBPSBarBottom = '\000r\000\003',
	MicrosoftWordMBPSBarFloating = '\000r\000\004',
	MicrosoftWordMBPSBarPopUp = '\000r\000\005',
	MicrosoftWordMBPSBarMenu = '\000r\000\006'
};
typedef enum MicrosoftWordMBPS MicrosoftWordMBPS;

enum MicrosoftWordMBPt {
	MicrosoftWordMBPtNoProtection = '\000s\000\000',
	MicrosoftWordMBPtNoCustomize = '\000s\000\001',
	MicrosoftWordMBPtNoResize = '\000s\000\002',
	MicrosoftWordMBPtNoMove = '\000s\000\004',
	MicrosoftWordMBPtNoChangeVisible = '\000s\000\010',
	MicrosoftWordMBPtNoChangeDock = '\000s\000\020',
	MicrosoftWordMBPtNoVerticalDock = '\000s\000 ',
	MicrosoftWordMBPtNoHorizontalDock = '\000s\000@'
};
typedef enum MicrosoftWordMBPt MicrosoftWordMBPt;

enum MicrosoftWordMBTY {
	MicrosoftWordMBTYNormalCommandBar = '\000t\000\000',
	MicrosoftWordMBTYMenubarCommandBar = '\000t\000\001',
	MicrosoftWordMBTYPopupCommandBar = '\000t\000\002'
};
typedef enum MicrosoftWordMBTY MicrosoftWordMBTY;

enum MicrosoftWordMCLT {
	MicrosoftWordMCLTControlCustom = '\000u\000\000',
	MicrosoftWordMCLTControlButton = '\000u\000\001',
	MicrosoftWordMCLTControlEdit = '\000u\000\002',
	MicrosoftWordMCLTControlDropDown = '\000u\000\003',
	MicrosoftWordMCLTControlCombobox = '\000u\000\004',
	MicrosoftWordMCLTButtonDropDown = '\000u\000\005',
	MicrosoftWordMCLTSplitDropDown = '\000u\000\006',
	MicrosoftWordMCLTOCXDropDown = '\000u\000\007',
	MicrosoftWordMCLTGenericDropDown = '\000u\000\010',
	MicrosoftWordMCLTGraphicDropDown = '\000u\000\011',
	MicrosoftWordMCLTControlPopup = '\000u\000\012',
	MicrosoftWordMCLTGraphicPopup = '\000u\000\013',
	MicrosoftWordMCLTButtonPopup = '\000u\000\014',
	MicrosoftWordMCLTSplitButtonPopup = '\000u\000\015',
	MicrosoftWordMCLTSplitButtonMRUPopup = '\000u\000\016',
	MicrosoftWordMCLTControlLabel = '\000u\000\017',
	MicrosoftWordMCLTExpandingGrid = '\000u\000\020',
	MicrosoftWordMCLTSplitExpandingGrid = '\000u\000\021',
	MicrosoftWordMCLTControlGrid = '\000u\000\022',
	MicrosoftWordMCLTControlGauge = '\000u\000\023',
	MicrosoftWordMCLTGraphicCombobox = '\000u\000\024',
	MicrosoftWordMCLTControlPane = '\000u\000\025',
	MicrosoftWordMCLTActiveX = '\000u\000\026',
	MicrosoftWordMCLTControlGroup = '\000u\000\027',
	MicrosoftWordMCLTControlTab = '\000u\000\030',
	MicrosoftWordMCLTControlSpinner = '\000u\000\031'
};
typedef enum MicrosoftWordMCLT MicrosoftWordMCLT;

enum MicrosoftWordMBns {
	MicrosoftWordMBnsButtonStateUp = '\000v\000\000',
	MicrosoftWordMBnsButtonStateDown = '\000u\377\377',
	MicrosoftWordMBnsButtonStateUnset = '\000v\000\002'
};
typedef enum MicrosoftWordMBns MicrosoftWordMBns;

enum MicrosoftWordMcOu {
	MicrosoftWordMcOuNeither = '\000w\000\000',
	MicrosoftWordMcOuServer = '\000w\000\001',
	MicrosoftWordMcOuClient = '\000w\000\002',
	MicrosoftWordMcOuBoth = '\000w\000\003'
};
typedef enum MicrosoftWordMcOu MicrosoftWordMcOu;

enum MicrosoftWordMBTs {
	MicrosoftWordMBTsButtonAutomatic = '\000x\000\000',
	MicrosoftWordMBTsButtonIcon = '\000x\000\001',
	MicrosoftWordMBTsButtonCaption = '\000x\000\002',
	MicrosoftWordMBTsButtonIconAndCaption = '\000x\000\003'
};
typedef enum MicrosoftWordMBTs MicrosoftWordMBTs;

enum MicrosoftWordMXcb {
	MicrosoftWordMXcbComboboxStyleNormal = '\000y\000\000',
	MicrosoftWordMXcbComboboxStyleLabel = '\000y\000\001'
};
typedef enum MicrosoftWordMXcb MicrosoftWordMXcb;

enum MicrosoftWordMMNA {
	MicrosoftWordMMNANone = '\000{\000\000',
	MicrosoftWordMMNARandom = '\000{\000\001',
	MicrosoftWordMMNAUnfold = '\000{\000\002',
	MicrosoftWordMMNASlide = '\000{\000\003'
};
typedef enum MicrosoftWordMMNA MicrosoftWordMMNA;

enum MicrosoftWordMHlT {
	MicrosoftWordMHlTHyperlinkTypeTextRange = '\000\226\000\000',
	MicrosoftWordMHlTHyperlinkTypeShape = '\000\226\000\001',
	MicrosoftWordMHlTHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum MicrosoftWordMHlT MicrosoftWordMHlT;

enum MicrosoftWordMXiM {
	MicrosoftWordMXiMAppendString = '\000\256\000\000',
	MicrosoftWordMXiMPostString = '\000\256\000\001'
};
typedef enum MicrosoftWordMXiM MicrosoftWordMXiM;

enum MicrosoftWordMANT {
	MicrosoftWordMANTIdle = '\000|\000\001',
	MicrosoftWordMANTGreeting = '\000|\000\002',
	MicrosoftWordMANTGoodbye = '\000|\000\003',
	MicrosoftWordMANTBeginSpeaking = '\000|\000\004',
	MicrosoftWordMANTCharacterSuccessMajor = '\000|\000\006',
	MicrosoftWordMANTGetAttentionMajor = '\000|\000\013',
	MicrosoftWordMANTGetAttentionMinor = '\000|\000\014',
	MicrosoftWordMANTSearching = '\000|\000\015',
	MicrosoftWordMANTPrinting = '\000|\000\022',
	MicrosoftWordMANTGestureRight = '\000|\000\023',
	MicrosoftWordMANTWritingNotingSomething = '\000|\000\026',
	MicrosoftWordMANTWorkingAtSomething = '\000|\000\027',
	MicrosoftWordMANTThinking = '\000|\000\030',
	MicrosoftWordMANTSendingMail = '\000|\000\031',
	MicrosoftWordMANTListensToComputer = '\000|\000\032',
	MicrosoftWordMANTDisappear = '\000|\000\037',
	MicrosoftWordMANTAppear = '\000|\000 ',
	MicrosoftWordMANTGetArtsy = '\000|\000d',
	MicrosoftWordMANTGetTechy = '\000|\000e',
	MicrosoftWordMANTGetWizardy = '\000|\000f',
	MicrosoftWordMANTCheckingSomething = '\000|\000g',
	MicrosoftWordMANTLookDown = '\000|\000h',
	MicrosoftWordMANTLookDownLeft = '\000|\000i',
	MicrosoftWordMANTLookDownRight = '\000|\000j',
	MicrosoftWordMANTLookLeft = '\000|\000k',
	MicrosoftWordMANTLookRight = '\000|\000l',
	MicrosoftWordMANTLookUp = '\000|\000m',
	MicrosoftWordMANTLookUpLeft = '\000|\000n',
	MicrosoftWordMANTLookUpRight = '\000|\000o',
	MicrosoftWordMANTSaving = '\000|\000p',
	MicrosoftWordMANTGestureDown = '\000|\000q',
	MicrosoftWordMANTGestureLeft = '\000|\000r',
	MicrosoftWordMANTGestureUp = '\000|\000s',
	MicrosoftWordMANTEmptyTrash = '\000|\000t'
};
typedef enum MicrosoftWordMANT MicrosoftWordMANT;

enum MicrosoftWordMBSt {
	MicrosoftWordMBStButtonNone = '\000}\000\000',
	MicrosoftWordMBStButtonOk = '\000}\000\001',
	MicrosoftWordMBStButtonCancel = '\000}\000\002',
	MicrosoftWordMBStButtonsOkCancel = '\000}\000\003',
	MicrosoftWordMBStButtonsYesNo = '\000}\000\004',
	MicrosoftWordMBStButtonsYesNoCancel = '\000}\000\005',
	MicrosoftWordMBStButtonsBackClose = '\000}\000\006',
	MicrosoftWordMBStButtonsNextClose = '\000}\000\007',
	MicrosoftWordMBStButtonsBackNextClose = '\000}\000\010',
	MicrosoftWordMBStButtonsRetryCancel = '\000}\000\011',
	MicrosoftWordMBStButtonsAbortRetryIgnore = '\000}\000\012',
	MicrosoftWordMBStButtonsSearchClose = '\000}\000\013',
	MicrosoftWordMBStButtonsBackNextSnooze = '\000}\000\014',
	MicrosoftWordMBStButtonsTipsOptionsClose = '\000}\000\015',
	MicrosoftWordMBStButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum MicrosoftWordMBSt MicrosoftWordMBSt;

enum MicrosoftWordMIct {
	MicrosoftWordMIctIconNone = '\000~\000\000',
	MicrosoftWordMIctIconApplication = '\000~\000\001',
	MicrosoftWordMIctIconAlert = '\000~\000\002',
	MicrosoftWordMIctIconTip = '\000~\000\003',
	MicrosoftWordMIctIconAlertCritical = '\000~\000e',
	MicrosoftWordMIctIconAlertWarning = '\000~\000g',
	MicrosoftWordMIctIconAlertInfo = '\000~\000h'
};
typedef enum MicrosoftWordMIct MicrosoftWordMIct;

enum MicrosoftWordMWAt {
	MicrosoftWordMWAtInactive = '\000\202\000\000',
	MicrosoftWordMWAtActive = '\000\202\000\001',
	MicrosoftWordMWAtSuspend = '\000\202\000\002',
	MicrosoftWordMWAtResume = '\000\202\000\003'
};
typedef enum MicrosoftWordMWAt MicrosoftWordMWAt;

enum MicrosoftWordMeDP {
	MicrosoftWordMeDPPropertyTypeNumber = '\000\242\000\001',
	MicrosoftWordMeDPPropertyTypeBoolean = '\000\242\000\002',
	MicrosoftWordMeDPPropertyTypeDate = '\000\242\000\003',
	MicrosoftWordMeDPPropertyTypeString = '\000\242\000\004',
	MicrosoftWordMeDPPropertyTypeFloat = '\000\242\000\005'
};
typedef enum MicrosoftWordMeDP MicrosoftWordMeDP;

enum MicrosoftWordMASc {
	MicrosoftWordMAScMsoAutomationSecurityLow = '\000\243\000\001',
	MicrosoftWordMAScMsoAutomationSecurityByUI = '\000\243\000\002',
	MicrosoftWordMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum MicrosoftWordMASc MicrosoftWordMASc;

enum MicrosoftWordMSsz {
	MicrosoftWordMSszResolution544x376 = '\000\204\000\000',
	MicrosoftWordMSszResolution640x480 = '\000\204\000\001',
	MicrosoftWordMSszResolution720x512 = '\000\204\000\002',
	MicrosoftWordMSszResolution800x600 = '\000\204\000\003',
	MicrosoftWordMSszResolution1024x768 = '\000\204\000\004',
	MicrosoftWordMSszResolution1152x882 = '\000\204\000\005',
	MicrosoftWordMSszResolution1152x900 = '\000\204\000\006',
	MicrosoftWordMSszResolution1280x1024 = '\000\204\000\007',
	MicrosoftWordMSszResolution1600x1200 = '\000\204\000\010',
	MicrosoftWordMSszResolution1800x1440 = '\000\204\000\011',
	MicrosoftWordMSszResolution1920x1200 = '\000\204\000\012'
};
typedef enum MicrosoftWordMSsz MicrosoftWordMSsz;

enum MicrosoftWordMChS {
	MicrosoftWordMChSArabicCharacterSet = '\000\205\000\001',
	MicrosoftWordMChSCyrillicCharacterSet = '\000\205\000\002',
	MicrosoftWordMChSEnglishCharacterSet = '\000\205\000\003',
	MicrosoftWordMChSGreekCharacterSet = '\000\205\000\004',
	MicrosoftWordMChSHebrewCharacterSet = '\000\205\000\005',
	MicrosoftWordMChSJapaneseCharacterSet = '\000\205\000\006',
	MicrosoftWordMChSKoreanCharacterSet = '\000\205\000\007',
	MicrosoftWordMChSMultilingualUnicodeCharacterSet = '\000\205\000\010',
	MicrosoftWordMChSSimplifiedChineseCharacterSet = '\000\205\000\011',
	MicrosoftWordMChSThaiCharacterSet = '\000\205\000\012',
	MicrosoftWordMChSTraditionalChineseCharacterSet = '\000\205\000\013',
	MicrosoftWordMChSVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum MicrosoftWordMChS MicrosoftWordMChS;

enum MicrosoftWordMtEn {
	MicrosoftWordMtEnEncodingThai = '\000\213\003j',
	MicrosoftWordMtEnEncodingJapaneseShiftJIS = '\000\213\003\244',
	MicrosoftWordMtEnEncodingSimplifiedChinese = '\000\213\003\250',
	MicrosoftWordMtEnEncodingKorean = '\000\213\003\265',
	MicrosoftWordMtEnEncodingBig5TraditionalChinese = '\000\213\003\266',
	MicrosoftWordMtEnEncodingLittleEndian = '\000\213\004\260',
	MicrosoftWordMtEnEncodingBigEndian = '\000\213\004\261',
	MicrosoftWordMtEnEncodingCentralEuropean = '\000\213\004\342',
	MicrosoftWordMtEnEncodingCyrillic = '\000\213\004\343',
	MicrosoftWordMtEnEncodingWestern = '\000\213\004\344',
	MicrosoftWordMtEnEncodingGreek = '\000\213\004\345',
	MicrosoftWordMtEnEncodingTurkish = '\000\213\004\346',
	MicrosoftWordMtEnEncodingHebrew = '\000\213\004\347',
	MicrosoftWordMtEnEncodingArabic = '\000\213\004\350',
	MicrosoftWordMtEnEncodingBaltic = '\000\213\004\351',
	MicrosoftWordMtEnEncodingVietnamese = '\000\213\004\352',
	MicrosoftWordMtEnEncodingISO88591Latin1 = '\000\213o\257',
	MicrosoftWordMtEnEncodingISO88592CentralEurope = '\000\213o\260',
	MicrosoftWordMtEnEncodingISO88593Latin3 = '\000\213o\261',
	MicrosoftWordMtEnEncodingISO88594Baltic = '\000\213o\262',
	MicrosoftWordMtEnEncodingISO88595Cyrillic = '\000\213o\263',
	MicrosoftWordMtEnEncodingISO88596Arabic = '\000\213o\264',
	MicrosoftWordMtEnEncodingISO88597Greek = '\000\213o\265',
	MicrosoftWordMtEnEncodingISO88598Hebrew = '\000\213o\266',
	MicrosoftWordMtEnEncodingISO88599Turkish = '\000\213o\267',
	MicrosoftWordMtEnEncodingISO885915Latin9 = '\000\213o\275',
	MicrosoftWordMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	MicrosoftWordMtEnEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	MicrosoftWordMtEnEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	MicrosoftWordMtEnEncodingISO2022KR = '\000\213\3041',
	MicrosoftWordMtEnEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	MicrosoftWordMtEnEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	MicrosoftWordMtEnEncodingMacRoman = '\000\213\'\020',
	MicrosoftWordMtEnEncodingMacJapanese = '\000\213\'\021',
	MicrosoftWordMtEnEncodingMacTraditionalChinese = '\000\213\'\022',
	MicrosoftWordMtEnEncodingMacKorean = '\000\213\'\023',
	MicrosoftWordMtEnEncodingMacArabic = '\000\213\'\024',
	MicrosoftWordMtEnEncodingMacHebrew = '\000\213\'\025',
	MicrosoftWordMtEnEncodingMacGreek1 = '\000\213\'\026',
	MicrosoftWordMtEnEncodingMacCyrillic = '\000\213\'\027',
	MicrosoftWordMtEnEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	MicrosoftWordMtEnEncodingMacRomania = '\000\213\'\032',
	MicrosoftWordMtEnEncodingMacUkraine = '\000\213\'!',
	MicrosoftWordMtEnEncodingMacLatin2 = '\000\213\'-',
	MicrosoftWordMtEnEncodingMacIcelandic = '\000\213\'_',
	MicrosoftWordMtEnEncodingMacTurkish = '\000\213\'a',
	MicrosoftWordMtEnEncodingMacCroatia = '\000\213\'b',
	MicrosoftWordMtEnEncodingEBCDICUSCanada = '\000\213\000%',
	MicrosoftWordMtEnEncodingEBCDICInternational = '\000\213\001\364',
	MicrosoftWordMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	MicrosoftWordMtEnEncodingEBCDICGreekModern = '\000\213\003k',
	MicrosoftWordMtEnEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	MicrosoftWordMtEnEncodingEBCDICGermany = '\000\213O1',
	MicrosoftWordMtEnEncodingEBCDICDenmarkNorway = '\000\213O5',
	MicrosoftWordMtEnEncodingEBCDICFinlandSweden = '\000\213O6',
	MicrosoftWordMtEnEncodingEBCDICItaly = '\000\213O8',
	MicrosoftWordMtEnEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	MicrosoftWordMtEnEncodingEBCDICUnitedKingdom = '\000\213O=',
	MicrosoftWordMtEnEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	MicrosoftWordMtEnEncodingEBCDICFrance = '\000\213OI',
	MicrosoftWordMtEnEncodingEBCDICArabic = '\000\213O\304',
	MicrosoftWordMtEnEncodingEBCDICGreek = '\000\213O\307',
	MicrosoftWordMtEnEncodingEBCDICHebrew = '\000\213O\310',
	MicrosoftWordMtEnEncodingEBCDICKoreanExtended = '\000\213Qa',
	MicrosoftWordMtEnEncodingEBCDICThai = '\000\213Qf',
	MicrosoftWordMtEnEncodingEBCDICIcelandic = '\000\213Q\207',
	MicrosoftWordMtEnEncodingEBCDICTurkish = '\000\213Q\251',
	MicrosoftWordMtEnEncodingEBCDICRussian = '\000\213Q\220',
	MicrosoftWordMtEnEncodingEBCDICSerbianBulgarian = '\000\213R!',
	MicrosoftWordMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	MicrosoftWordMtEnEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	MicrosoftWordMtEnEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	MicrosoftWordMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	MicrosoftWordMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	MicrosoftWordMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	MicrosoftWordMtEnEncodingOEMUnitedStates = '\000\213\001\265',
	MicrosoftWordMtEnEncodingOEMGreek = '\000\213\002\341',
	MicrosoftWordMtEnEncodingOEMBaltic = '\000\213\003\007',
	MicrosoftWordMtEnEncodingOEMMultilingualLatinI = '\000\213\003R',
	MicrosoftWordMtEnEncodingOEMMultilingualLatinII = '\000\213\003T',
	MicrosoftWordMtEnEncodingOEMCyrillic = '\000\213\003W',
	MicrosoftWordMtEnEncodingOEMTurkish = '\000\213\003Y',
	MicrosoftWordMtEnEncodingOEMPortuguese = '\000\213\003\\',
	MicrosoftWordMtEnEncodingOEMIcelandic = '\000\213\003]',
	MicrosoftWordMtEnEncodingOEMHebrew = '\000\213\003^',
	MicrosoftWordMtEnEncodingOEMCanadianFrench = '\000\213\003_',
	MicrosoftWordMtEnEncodingOEMArabic = '\000\213\003`',
	MicrosoftWordMtEnEncodingOEMNordic = '\000\213\003a',
	MicrosoftWordMtEnEncodingOEMCyrillicII = '\000\213\003b',
	MicrosoftWordMtEnEncodingOEMModernGreek = '\000\213\003e',
	MicrosoftWordMtEnEncodingEUCJapanese = '\000\213\312\334',
	MicrosoftWordMtEnEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	MicrosoftWordMtEnEncodingEUCKorean = '\000\213\312\355',
	MicrosoftWordMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	MicrosoftWordMtEnEncodingDevanagari = '\000\213\336\252',
	MicrosoftWordMtEnEncodingBengali = '\000\213\336\253',
	MicrosoftWordMtEnEncodingTamil = '\000\213\336\254',
	MicrosoftWordMtEnEncodingTelugu = '\000\213\336\255',
	MicrosoftWordMtEnEncodingAssamese = '\000\213\336\256',
	MicrosoftWordMtEnEncodingOriya = '\000\213\336\257',
	MicrosoftWordMtEnEncodingKannada = '\000\213\336\260',
	MicrosoftWordMtEnEncodingMalayalam = '\000\213\336\261',
	MicrosoftWordMtEnEncodingGujarati = '\000\213\336\262',
	MicrosoftWordMtEnEncodingPunjabi = '\000\213\336\263',
	MicrosoftWordMtEnEncodingArabicASMO = '\000\213\002\304',
	MicrosoftWordMtEnEncodingArabicTransparentASMO = '\000\213\002\320',
	MicrosoftWordMtEnEncodingKoreanJohab = '\000\213\005Q',
	MicrosoftWordMtEnEncodingTaiwanCNS = '\000\213N ',
	MicrosoftWordMtEnEncodingTaiwanTCA = '\000\213N!',
	MicrosoftWordMtEnEncodingTaiwanEten = '\000\213N\"',
	MicrosoftWordMtEnEncodingTaiwanIBM5550 = '\000\213N#',
	MicrosoftWordMtEnEncodingTaiwanTeletext = '\000\213N$',
	MicrosoftWordMtEnEncodingTaiwanWang = '\000\213N%',
	MicrosoftWordMtEnEncodingIA5IRV = '\000\213N\211',
	MicrosoftWordMtEnEncodingIA5German = '\000\213N\212',
	MicrosoftWordMtEnEncodingIA5Swedish = '\000\213N\213',
	MicrosoftWordMtEnEncodingIA5Norwegian = '\000\213N\214',
	MicrosoftWordMtEnEncodingUSASCII = '\000\213N\237',
	MicrosoftWordMtEnEncodingT61 = '\000\213O%',
	MicrosoftWordMtEnEncodingISO6937NonspacingAccent = '\000\213O-',
	MicrosoftWordMtEnEncodingKOI8R = '\000\213Q\202',
	MicrosoftWordMtEnEncodingExtAlphaLowercase = '\000\213R#',
	MicrosoftWordMtEnEncodingKOI8U = '\000\213Uj',
	MicrosoftWordMtEnEncodingEuropa3 = '\000\213qI',
	MicrosoftWordMtEnEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	MicrosoftWordMtEnEncodingUTF7 = '\000\213\375\350',
	MicrosoftWordMtEnEncodingUTF8 = '\000\213\375\351'
};
typedef enum MicrosoftWordMtEn MicrosoftWordMtEn;

enum MicrosoftWord4000 {
	MicrosoftWord4000CommandBar = 'msCB',
	MicrosoftWord4000CommandBarControl = 'mCBC'
};
typedef enum MicrosoftWord4000 MicrosoftWord4000;

enum MicrosoftWordEMvS {
	MicrosoftWordEMvSTowardsTheEnd = '\002\266\000\000',
	MicrosoftWordEMvSTowardsTheStart = '\002\266\000\001',
	MicrosoftWordEMvSTowardsTheLeft = '\002\266\000\002',
	MicrosoftWordEMvSTowardsTheRight = '\002\266\000\003',
	MicrosoftWordEMvSTowardsTheTop = '\002\266\000\004',
	MicrosoftWordEMvSTowardsTheBottom = '\002\266\000\005'
};
typedef enum MicrosoftWordEMvS MicrosoftWordEMvS;

enum MicrosoftWordEMUW {
	MicrosoftWordEMUWTowardsTheTail = '\002\267\000\000',
	MicrosoftWordEMUWTowardsTheBeginning = '\002\267\000\001'
};
typedef enum MicrosoftWordEMUW MicrosoftWordEMUW;

enum MicrosoftWordEiAb {
	MicrosoftWordEiAbAbove = '\002\270\000\000',
	MicrosoftWordEiAbBelow = '\002\270\000\001'
};
typedef enum MicrosoftWordEiAb MicrosoftWordEiAb;

enum MicrosoftWordEiBa {
	MicrosoftWordEiBaBeforeTheObject = '\002\271\000\000',
	MicrosoftWordEiBaAfterTheObject = '\002\271\000\001'
};
typedef enum MicrosoftWordEiBa MicrosoftWordEiBa;

enum MicrosoftWordEiRL {
	MicrosoftWordEiRLInsertOnTheRight = '\002\273\000\000',
	MicrosoftWordEiRLInsertOnTheLeft = '\002\273\000\001'
};
typedef enum MicrosoftWordEiRL MicrosoftWordEiRL;

enum MicrosoftWordEIPt {
	MicrosoftWordEIPtOffset = '\003\037\000\000',
	MicrosoftWordEIPtLocationReference = '\003\037\000\001'
};
typedef enum MicrosoftWordEIPt MicrosoftWordEIPt;

enum MicrosoftWordEFRt {
	MicrosoftWordEFRtTextRange = '\003\036\000\000',
	MicrosoftWordEFRtInsertionPoint = '\003\036\000\001'
};
typedef enum MicrosoftWordEFRt MicrosoftWordEFRt;

enum MicrosoftWordE101 {
	MicrosoftWordE101TypeNormalTemplate = '\001\365\000\000',
	MicrosoftWordE101TypeGlobalTemplate = '\001\365\000\001',
	MicrosoftWordE101TypeAttachedTemplate = '\001\365\000\002'
};
typedef enum MicrosoftWordE101 MicrosoftWordE101;

enum MicrosoftWordE102 {
	MicrosoftWordE102ContinueDisabled = '\001\366\000\000',
	MicrosoftWordE102ResetList = '\001\366\000\001',
	MicrosoftWordE102ContinueList = '\001\366\000\002'
};
typedef enum MicrosoftWordE102 MicrosoftWordE102;

enum MicrosoftWordE103 {
	MicrosoftWordE103IMEModeNoControl = '\001\367\000\000',
	MicrosoftWordE103IMEModeOn = '\001\367\000\001',
	MicrosoftWordE103IMEModeOff = '\001\367\000\002',
	MicrosoftWordE103IMEModeHiragana = '\001\367\000\004',
	MicrosoftWordE103IMEModeKatakana = '\001\367\000\005',
	MicrosoftWordE103IMEModeKatakanaHalf = '\001\367\000\006',
	MicrosoftWordE103IMEModeAlphaFull = '\001\367\000\007',
	MicrosoftWordE103IMEModeAlpha = '\001\367\000\010',
	MicrosoftWordE103IMEModeHangulFull = '\001\367\000\011',
	MicrosoftWordE103IMEModeHangul = '\001\367\000\012'
};
typedef enum MicrosoftWordE103 MicrosoftWordE103;

enum MicrosoftWordE104 {
	MicrosoftWordE104BaselineAlignTop = '\001\370\000\000',
	MicrosoftWordE104BaselineAlignCenter = '\001\370\000\001',
	MicrosoftWordE104BaselineAlignBaseline = '\001\370\000\002',
	MicrosoftWordE104BaselineAlignEastAsian50 = '\001\370\000\003',
	MicrosoftWordE104BaselineAlignAuto = '\001\370\000\004'
};
typedef enum MicrosoftWordE104 MicrosoftWordE104;

enum MicrosoftWordE105 {
	MicrosoftWordE105IndexFilterNone = '\001\371\000\000',
	MicrosoftWordE105IndexFilterAiueo = '\001\371\000\001',
	MicrosoftWordE105IndexFilterAkasatana = '\001\371\000\002',
	MicrosoftWordE105IndexFilterChosung = '\001\371\000\003',
	MicrosoftWordE105IndexFilterLow = '\001\371\000\004',
	MicrosoftWordE105IndexFilterMedium = '\001\371\000\005',
	MicrosoftWordE105IndexFilterFull = '\001\371\000\006'
};
typedef enum MicrosoftWordE105 MicrosoftWordE105;

enum MicrosoftWordE106 {
	MicrosoftWordE106IndexSortByStroke = '\001\372\000\000',
	MicrosoftWordE106IndexSortBySyllable = '\001\372\000\001'
};
typedef enum MicrosoftWordE106 MicrosoftWordE106;

enum MicrosoftWordE107 {
	MicrosoftWordE107JustificationModeExpand = '\001\373\000\000',
	MicrosoftWordE107JustificationModeCompress = '\001\373\000\001',
	MicrosoftWordE107JustificationModeCompressKana = '\001\373\000\002'
};
typedef enum MicrosoftWordE107 MicrosoftWordE107;

enum MicrosoftWordE108 {
	MicrosoftWordE108EastAsianLineBreakLevelNormal = '\001\374\000\000',
	MicrosoftWordE108EastAsianLineBreakLevelStrict = '\001\374\000\001',
	MicrosoftWordE108EastAsianLineBreakLevelCustom = '\001\374\000\002'
};
typedef enum MicrosoftWordE108 MicrosoftWordE108;

enum MicrosoftWordE109 {
	MicrosoftWordE109HangulToHanja = '\001\375\000\000',
	MicrosoftWordE109HanjaToHangul = '\001\375\000\001'
};
typedef enum MicrosoftWordE109 MicrosoftWordE109;

enum MicrosoftWordE110 {
	MicrosoftWordE110Auto = '\001\376\000\000',
	MicrosoftWordE110Black = '\001\376\000\001',
	MicrosoftWordE110Blue = '\001\376\000\002',
	MicrosoftWordE110Turquoise = '\001\376\000\003',
	MicrosoftWordE110BrightGreen = '\001\376\000\004',
	MicrosoftWordE110Pink = '\001\376\000\005',
	MicrosoftWordE110Red = '\001\376\000\006',
	MicrosoftWordE110Yellow = '\001\376\000\007',
	MicrosoftWordE110White = '\001\376\000\010',
	MicrosoftWordE110DarkBlue = '\001\376\000\011',
	MicrosoftWordE110Teal = '\001\376\000\012',
	MicrosoftWordE110Green = '\001\376\000\013',
	MicrosoftWordE110Violet = '\001\376\000\014',
	MicrosoftWordE110DarkRed = '\001\376\000\015',
	MicrosoftWordE110DarkYellow = '\001\376\000\016',
	MicrosoftWordE110Gray50 = '\001\376\000\017',
	MicrosoftWordE110Gray25 = '\001\376\000\020',
	MicrosoftWordE110ByAuthor = '\001\375\377\377',
	MicrosoftWordE110NoHighlight = '\001\376\000\000'
};
typedef enum MicrosoftWordE110 MicrosoftWordE110;

enum MicrosoftWordE111 {
	MicrosoftWordE111ColorAutomatic = '\377\000\000\000',
	MicrosoftWordE111ColorBlack = '\000\000\000\000',
	MicrosoftWordE111ColorBlue = '\000\377\000\000',
	MicrosoftWordE111ColorPink = '\000\377\000\377',
	MicrosoftWordE111ColorRed = '\000\000\000\377',
	MicrosoftWordE111ColorYellow = '\000\000\377\377',
	MicrosoftWordE111ColorTurquoise = '\000\377\377\000',
	MicrosoftWordE111ColorBrightGreen = '\000\000\377\000',
	MicrosoftWordE111ColorGreen = '\000\000\200\000',
	MicrosoftWordE111ColorWhite = '\000\377\377\377',
	MicrosoftWordE111ColorDarkBlue = '\000\200\000\000',
	MicrosoftWordE111ColorTeal = '\000\200\200\000',
	MicrosoftWordE111ColorViolet = '\000\200\000\200',
	MicrosoftWordE111ColorDarkGreen = '\000\0003\000',
	MicrosoftWordE111ColorDarkRed = '\000\000\000\200',
	MicrosoftWordE111ColorDarkYellow = '\000\000\200\200',
	MicrosoftWordE111ColorBrown = '\000\0003\231',
	MicrosoftWordE111ColorOliveGreen = '\000\00033',
	MicrosoftWordE111ColorDarkTeal = '\000f3\000',
	MicrosoftWordE111ColorIndigo = '\000\23133',
	MicrosoftWordE111ColorOrange = '\000\000f\377',
	MicrosoftWordE111ColorBlueGray = '\000\231ff',
	MicrosoftWordE111ColorLightOrange = '\000\000\231\377',
	MicrosoftWordE111ColorLime = '\000\000\314\231',
	MicrosoftWordE111ColorSeaGreen = '\000f\2313',
	MicrosoftWordE111ColorAqua = '\000\314\3143',
	MicrosoftWordE111ColorLightBlue = '\000\377f3',
	MicrosoftWordE111ColorGold = '\000\000\314\377',
	MicrosoftWordE111ColorSkyBlue = '\000\377\314\000',
	MicrosoftWordE111ColorPlum = '\000f3\231',
	MicrosoftWordE111ColorRose = '\000\314\231\377',
	MicrosoftWordE111ColorTan = '\000\231\314\377',
	MicrosoftWordE111ColorLightYellow = '\000\231\377\377',
	MicrosoftWordE111ColorLightGreen = '\000\314\377\314',
	MicrosoftWordE111ColorLightTurquoise = '\000\377\377\314',
	MicrosoftWordE111ColorPaleBlue = '\000\377\314\231',
	MicrosoftWordE111ColorLavender = '\000\377\231\314',
	MicrosoftWordE111ColorGray05 = '\000\363\363\363',
	MicrosoftWordE111ColorGray10 = '\000\346\346\346',
	MicrosoftWordE111ColorGray125 = '\000\340\340\340',
	MicrosoftWordE111ColorGray15 = '\000\331\331\331',
	MicrosoftWordE111ColorGray20 = '\000\314\314\314',
	MicrosoftWordE111ColorGray25 = '\000\300\300\300',
	MicrosoftWordE111ColorGray30 = '\000\263\263\263',
	MicrosoftWordE111ColorGray35 = '\000\246\246\246',
	MicrosoftWordE111ColorGray375 = '\000\240\240\240',
	MicrosoftWordE111ColorGray40 = '\000\231\231\231',
	MicrosoftWordE111ColorGray45 = '\000\214\214\214',
	MicrosoftWordE111ColorGray50 = '\000\200\200\200',
	MicrosoftWordE111ColorGray55 = '\000sss',
	MicrosoftWordE111ColorGray60 = '\000fff',
	MicrosoftWordE111ColorGray625 = '\000```',
	MicrosoftWordE111ColorGray65 = '\000YYY',
	MicrosoftWordE111ColorGray70 = '\000LLL',
	MicrosoftWordE111ColorGray75 = '\000@@@',
	MicrosoftWordE111ColorGray80 = '\000333',
	MicrosoftWordE111ColorGray85 = '\000&&&',
	MicrosoftWordE111ColorGray875 = '\000   ',
	MicrosoftWordE111ColorGray90 = '\000\031\031\031',
	MicrosoftWordE111ColorGray95 = '\000\014\014\014'
};
typedef enum MicrosoftWordE111 MicrosoftWordE111;

enum MicrosoftWordE112 {
	MicrosoftWordE112TextureNone = '\002\000\000\000',
	MicrosoftWordE112Texture2Pt5Percent = '\002\000\000\031',
	MicrosoftWordE112Texture5Percent = '\002\000\0002',
	MicrosoftWordE112Texture7Pt5Percent = '\002\000\000K',
	MicrosoftWordE112Texture10Percent = '\002\000\000d',
	MicrosoftWordE112Texture12Pt5Percent = '\002\000\000}',
	MicrosoftWordE112Texture15Percent = '\002\000\000\226',
	MicrosoftWordE112Texture17Pt5Percent = '\002\000\000\257',
	MicrosoftWordE112Texture20Percent = '\002\000\000\310',
	MicrosoftWordE112Texture22Pt5Percent = '\002\000\000\341',
	MicrosoftWordE112Texture25Percent = '\002\000\000\372',
	MicrosoftWordE112Texture27Pt5Percent = '\002\000\001\023',
	MicrosoftWordE112Texture30Percent = '\002\000\001,',
	MicrosoftWordE112Texture32Pt5Percent = '\002\000\001E',
	MicrosoftWordE112Texture35Percent = '\002\000\001^',
	MicrosoftWordE112Texture37Pt5Percent = '\002\000\001w',
	MicrosoftWordE112Texture40Percent = '\002\000\001\220',
	MicrosoftWordE112Texture42Pt5Percent = '\002\000\001\251',
	MicrosoftWordE112Texture45Percent = '\002\000\001\302',
	MicrosoftWordE112Texture47Pt5Percent = '\002\000\001\333',
	MicrosoftWordE112Texture50Percent = '\002\000\001\364',
	MicrosoftWordE112Texture52Pt5Percent = '\002\000\002\015',
	MicrosoftWordE112Texture55Percent = '\002\000\002&',
	MicrosoftWordE112Texture57Pt5Percent = '\002\000\002\?',
	MicrosoftWordE112Texture60Percent = '\002\000\002X',
	MicrosoftWordE112Texture62Pt5Percent = '\002\000\002q',
	MicrosoftWordE112Texture65Percent = '\002\000\002\212',
	MicrosoftWordE112Texture67Pt5Percent = '\002\000\002\243',
	MicrosoftWordE112Texture70Percent = '\002\000\002\274',
	MicrosoftWordE112Texture72Pt5Percent = '\002\000\002\325',
	MicrosoftWordE112Texture75Percent = '\002\000\002\356',
	MicrosoftWordE112Texture77Pt5Percent = '\002\000\003\007',
	MicrosoftWordE112Texture80Percent = '\002\000\003 ',
	MicrosoftWordE112Texture82Pt5Percent = '\002\000\0039',
	MicrosoftWordE112Texture85Percent = '\002\000\003R',
	MicrosoftWordE112Texture87Pt5Percent = '\002\000\003k',
	MicrosoftWordE112Texture90Percent = '\002\000\003\204',
	MicrosoftWordE112Texture92Pt5Percent = '\002\000\003\235',
	MicrosoftWordE112Texture95Percent = '\002\000\003\266',
	MicrosoftWordE112Texture97Pt5Percent = '\002\000\003\317',
	MicrosoftWordE112TextureSolid = '\002\000\003\350',
	MicrosoftWordE112TextureDarkHorizontal = '\001\377\377\377',
	MicrosoftWordE112TextureDarkVertical = '\001\377\377\376',
	MicrosoftWordE112TextureDarkDiagonalDown = '\001\377\377\375',
	MicrosoftWordE112TextureDarkDiagonalUp = '\001\377\377\374',
	MicrosoftWordE112TextureDarkCross = '\001\377\377\373',
	MicrosoftWordE112TextureDarkDiagonalCross = '\001\377\377\372',
	MicrosoftWordE112TextureHorizontal = '\001\377\377\371',
	MicrosoftWordE112TextureVertical = '\001\377\377\370',
	MicrosoftWordE112TextureDiagonalDown = '\001\377\377\367',
	MicrosoftWordE112TextureDiagonalUp = '\001\377\377\366',
	MicrosoftWordE112TextureCross = '\001\377\377\365',
	MicrosoftWordE112TextureDiagonalCross = '\001\377\377\364'
};
typedef enum MicrosoftWordE112 MicrosoftWordE112;

enum MicrosoftWordE113 {
	MicrosoftWordE113UnderlineNone = '\002\001\000\000',
	MicrosoftWordE113UnderlineSingle = '\002\001\000\001',
	MicrosoftWordE113UnderlineWords = '\002\001\000\002',
	MicrosoftWordE113UnderlineDouble = '\002\001\000\003',
	MicrosoftWordE113UnderlineDotted = '\002\001\000\004',
	MicrosoftWordE113UnderlineThick = '\002\001\000\006',
	MicrosoftWordE113UnderlineDash = '\002\001\000\007',
	MicrosoftWordE113UnderlineDotDash = '\002\001\000\011',
	MicrosoftWordE113UnderlineDotDotDash = '\002\001\000\012',
	MicrosoftWordE113UnderlineWavy = '\002\001\000\013',
	MicrosoftWordE113UnderlineWavyHeavy = '\002\001\000\033',
	MicrosoftWordE113UnderlineDottedHeavy = '\002\001\000\024',
	MicrosoftWordE113UnderlineDashHeavy = '\002\001\000\027',
	MicrosoftWordE113UnderlineDotDashHeavy = '\002\001\000\031',
	MicrosoftWordE113UnderlineDotDotDashHeavy = '\002\001\000\032',
	MicrosoftWordE113UnderlineDashLong = '\002\001\000\'',
	MicrosoftWordE113UnderlineDashLongHeavy = '\002\001\0007',
	MicrosoftWordE113UnderlineWavyDouble = '\002\001\000+'
};
typedef enum MicrosoftWordE113 MicrosoftWordE113;

enum MicrosoftWordE114 {
	MicrosoftWordE114EmphasisMarkNone = '\002\002\000\000',
	MicrosoftWordE114EmphasisMarkOverSolidCircle = '\002\002\000\001',
	MicrosoftWordE114EmphasisMarkOverComma = '\002\002\000\002',
	MicrosoftWordE114EmphasisMarkOverWhiteCircle = '\002\002\000\003',
	MicrosoftWordE114EmphasisMarkUnderSolidCircle = '\002\002\000\004'
};
typedef enum MicrosoftWordE114 MicrosoftWordE114;

enum MicrosoftWordE115 {
	MicrosoftWordE115ListSeparator = '\002\003\000\021',
	MicrosoftWordE115DecimalSeparator = '\002\003\000\022',
	MicrosoftWordE115ThousandsSeparator = '\002\003\000\023',
	MicrosoftWordE115CurrencyCode = '\002\003\000\024',
	MicrosoftWordE115TwentyFourHourClock = '\002\003\000\025',
	MicrosoftWordE115InternationalAm = '\002\003\000\026',
	MicrosoftWordE115InternationalPm = '\002\003\000\027',
	MicrosoftWordE115TimeSeparator = '\002\003\000\030',
	MicrosoftWordE115DateSeparator = '\002\003\000\031',
	MicrosoftWordE115ProductLanguageID = '\002\003\000\032'
};
typedef enum MicrosoftWordE115 MicrosoftWordE115;

enum MicrosoftWordE116 {
	MicrosoftWordE116AutoExec = '\002\004\000\000',
	MicrosoftWordE116AutoNew = '\002\004\000\001',
	MicrosoftWordE116AutoOpen = '\002\004\000\002',
	MicrosoftWordE116AutoClose = '\002\004\000\003',
	MicrosoftWordE116AutoExit = '\002\004\000\004'
};
typedef enum MicrosoftWordE116 MicrosoftWordE116;

enum MicrosoftWordE117 {
	MicrosoftWordE117CaptionPositionAbove = '\002\005\000\000',
	MicrosoftWordE117CaptionPositionBelow = '\002\005\000\001'
};
typedef enum MicrosoftWordE117 MicrosoftWordE117;

enum MicrosoftWordE118 {
	MicrosoftWordE118Us = '\002\006\000\001',
	MicrosoftWordE118Canada = '\002\006\000\002',
	MicrosoftWordE118LatinAmerica = '\002\006\000\003',
	MicrosoftWordE118Netherlands = '\002\006\000\037',
	MicrosoftWordE118France = '\002\006\000!',
	MicrosoftWordE118Spain = '\002\006\000\"',
	MicrosoftWordE118Italy = '\002\006\000\'',
	MicrosoftWordE118Uk = '\002\006\000,',
	MicrosoftWordE118Denmark = '\002\006\000-',
	MicrosoftWordE118Sweden = '\002\006\000.',
	MicrosoftWordE118Norway = '\002\006\000/',
	MicrosoftWordE118Germany = '\002\006\0001',
	MicrosoftWordE118Peru = '\002\006\0003',
	MicrosoftWordE118Mexico = '\002\006\0004',
	MicrosoftWordE118Argentina = '\002\006\0006',
	MicrosoftWordE118Brazil = '\002\006\0007',
	MicrosoftWordE118Chile = '\002\006\0008',
	MicrosoftWordE118Venezuela = '\002\006\000:',
	MicrosoftWordE118Japan = '\002\006\000Q',
	MicrosoftWordE118Taiwan = '\002\006\003v',
	MicrosoftWordE118China = '\002\006\000V',
	MicrosoftWordE118Korea = '\002\006\000R',
	MicrosoftWordE118Finland = '\002\006\001f',
	MicrosoftWordE118Iceland = '\002\006\001b'
};
typedef enum MicrosoftWordE118 MicrosoftWordE118;

enum MicrosoftWordE119 {
	MicrosoftWordE119HeadingSeparatorNone = '\002\007\000\000',
	MicrosoftWordE119HeadingSeparatorBlankLine = '\002\007\000\001',
	MicrosoftWordE119HeadingSeparatorLetter = '\002\007\000\002',
	MicrosoftWordE119HeadingSeparatorLetterLow = '\002\007\000\003',
	MicrosoftWordE119HeadingSeparatorLetterFull = '\002\007\000\004'
};
typedef enum MicrosoftWordE119 MicrosoftWordE119;

enum MicrosoftWordE120 {
	MicrosoftWordE120SeparatorHyphen = '\002\010\000\000',
	MicrosoftWordE120SeparatorPeriod = '\002\010\000\001',
	MicrosoftWordE120SeparatorColon = '\002\010\000\002',
	MicrosoftWordE120SeparatorEmDash = '\002\010\000\003',
	MicrosoftWordE120SeparatorEnDash = '\002\010\000\004'
};
typedef enum MicrosoftWordE120 MicrosoftWordE120;

enum MicrosoftWordE121 {
	MicrosoftWordE121AlignPageNumberLeft = '\002\011\000\000',
	MicrosoftWordE121AlignPageNumberCenter = '\002\011\000\001',
	MicrosoftWordE121AlignPageNumberRight = '\002\011\000\002',
	MicrosoftWordE121AlignPageNumberInside = '\002\011\000\003',
	MicrosoftWordE121AlignPageNumberOutside = '\002\011\000\004'
};
typedef enum MicrosoftWordE121 MicrosoftWordE121;

enum MicrosoftWordE122 {
	MicrosoftWordE122BorderTop = '\002\011\377\377',
	MicrosoftWordE122BorderLeft = '\002\011\377\376',
	MicrosoftWordE122BorderBottom = '\002\011\377\375',
	MicrosoftWordE122BorderRight = '\002\011\377\374',
	MicrosoftWordE122BorderHorizontal = '\002\011\377\373',
	MicrosoftWordE122BorderVertical = '\002\011\377\372',
	MicrosoftWordE122BorderDiagonalDown = '\002\011\377\371',
	MicrosoftWordE122BorderDiagonalUp = '\002\011\377\370'
};
typedef enum MicrosoftWordE122 MicrosoftWordE122;

enum MicrosoftWordE123 {
	MicrosoftWordE123FrameTop = '\001\373\275\301',
	MicrosoftWordE123FrameLeft = '\001\373\275\302',
	MicrosoftWordE123FrameBottom = '\001\373\275\303',
	MicrosoftWordE123FrameRight = '\001\373\275\304',
	MicrosoftWordE123FrameCenter = '\001\373\275\305',
	MicrosoftWordE123FrameInside = '\001\373\275\306',
	MicrosoftWordE123FrameOutside = '\001\373\275\307'
};
typedef enum MicrosoftWordE123 MicrosoftWordE123;

enum MicrosoftWordE124 {
	MicrosoftWordE124AnimationNone = '\002\014\000\000',
	MicrosoftWordE124AnimationLasVegasLights = '\002\014\000\001',
	MicrosoftWordE124AnimationBlinkingBackground = '\002\014\000\002',
	MicrosoftWordE124AnimationSparkleText = '\002\014\000\003',
	MicrosoftWordE124AnimationMarchingBlackAnts = '\002\014\000\004',
	MicrosoftWordE124AnimationMarchingRedAnts = '\002\014\000\005',
	MicrosoftWordE124AnimationShimmer = '\002\014\000\006'
};
typedef enum MicrosoftWordE124 MicrosoftWordE124;

enum MicrosoftWordE125 {
	MicrosoftWordE125NextCase = '\002\014\377\377',
	MicrosoftWordE125LowerCase = '\002\015\000\000',
	MicrosoftWordE125UpperCase = '\002\015\000\001',
	MicrosoftWordE125TitleWord = '\002\015\000\002',
	MicrosoftWordE125TitleSentence = '\002\015\000\004',
	MicrosoftWordE125ToggleCase = '\002\015\000\005',
	MicrosoftWordE125CaseHalfWidth = '\002\015\000\006',
	MicrosoftWordE125CaseFullWidth = '\002\015\000\007',
	MicrosoftWordE125CaseKatakana = '\002\015\000\010',
	MicrosoftWordE125CaseHiragana = '\002\015\000\011'
};
typedef enum MicrosoftWordE125 MicrosoftWordE125;

enum MicrosoftWordE127 {
	MicrosoftWordE12710Sentences = '\002\016\377\376',
	MicrosoftWordE12720Sentences = '\002\016\377\375',
	MicrosoftWordE127100Words = '\002\016\377\374',
	MicrosoftWordE127500Words = '\002\016\377\373',
	MicrosoftWordE12710Percent = '\002\016\377\372',
	MicrosoftWordE12725Percent = '\002\016\377\371',
	MicrosoftWordE12750Percent = '\002\016\377\370',
	MicrosoftWordE12775Percent = '\002\016\377\367'
};
typedef enum MicrosoftWordE127 MicrosoftWordE127;

enum MicrosoftWordE128 {
	MicrosoftWordE128StyleTypeParagraph = '\002\020\000\001',
	MicrosoftWordE128StyleTypeCharacter = '\002\020\000\002',
	MicrosoftWordE128StyleTypeTable = '\002\020\000\003',
	MicrosoftWordE128StyleTypeList = '\002\020\000\004'
};
typedef enum MicrosoftWordE128 MicrosoftWordE128;

enum MicrosoftWordE129 {
	MicrosoftWordE129ACharacterItem = '\002\021\000\001',
	MicrosoftWordE129AWordItem = '\002\021\000\002',
	MicrosoftWordE129ASentenceItem = '\002\021\000\003',
	MicrosoftWordE129AParagraphItem = '\002\021\000\004',
	MicrosoftWordE129ALineItem = '\002\021\000\005',
	MicrosoftWordE129AStoryItem = '\002\021\000\006',
	MicrosoftWordE129AScreen = '\002\021\000\007',
	MicrosoftWordE129ASection = '\002\021\000\010',
	MicrosoftWordE129AColumn = '\002\021\000\011',
	MicrosoftWordE129ARow = '\002\021\000\012',
	MicrosoftWordE129AWindow = '\002\021\000\013',
	MicrosoftWordE129ACell = '\002\021\000\014',
	MicrosoftWordE129ACharacterFormatting = '\002\021\000\015',
	MicrosoftWordE129AParagraphFormatting = '\002\021\000\016',
	MicrosoftWordE129ATable = '\002\021\000\017',
	MicrosoftWordE129AnItemUnit = '\002\021\000\020'
};
typedef enum MicrosoftWordE129 MicrosoftWordE129;

enum MicrosoftWordE130 {
	MicrosoftWordE130GotoABookmarkItem = '\002\021\377\377',
	MicrosoftWordE130GotoASectionItem = '\002\022\000\000',
	MicrosoftWordE130GotoAPageItem = '\002\022\000\001',
	MicrosoftWordE130GotoATableItem = '\002\022\000\002',
	MicrosoftWordE130GotoALineItem = '\002\022\000\003',
	MicrosoftWordE130GotoAFootnoteItem = '\002\022\000\004',
	MicrosoftWordE130GotoAnEndnoteItem = '\002\022\000\005',
	MicrosoftWordE130GotoACommentItem = '\002\022\000\006',
	MicrosoftWordE130GotoAFieldItem = '\002\022\000\007',
	MicrosoftWordE130GotoAGraphic = '\002\022\000\010',
	MicrosoftWordE130GotoAnObjectItem = '\002\022\000\011',
	MicrosoftWordE130GotoAnEquation = '\002\022\000\012',
	MicrosoftWordE130GotoAHeadingItem = '\002\022\000\013',
	MicrosoftWordE130GotoAPercent = '\002\022\000\014',
	MicrosoftWordE130GotoASpellingError = '\002\022\000\015',
	MicrosoftWordE130GotoAGrammaticalError = '\002\022\000\016',
	MicrosoftWordE130GotoAProofreadingError = '\002\022\000\017'
};
typedef enum MicrosoftWordE130 MicrosoftWordE130;

enum MicrosoftWordE131 {
	MicrosoftWordE131TheFirstItem = '\002\023\000\001',
	MicrosoftWordE131TheLastItem = '\002\022\377\377',
	MicrosoftWordE131TheNextItem = '\002\023\000\002',
	MicrosoftWordE131Relative = '\002\023\000\002',
	MicrosoftWordE131ThePreviousItem = '\002\023\000\003',
	MicrosoftWordE131Absolute = '\002\023\000\001'
};
typedef enum MicrosoftWordE131 MicrosoftWordE131;

enum MicrosoftWordE132 {
	MicrosoftWordE132CollapseStart = '\002\024\000\001',
	MicrosoftWordE132CollapseEnd = '\002\024\000\000'
};
typedef enum MicrosoftWordE132 MicrosoftWordE132;

enum MicrosoftWordE133 {
	MicrosoftWordE133RowHeightAuto = '\002\025\000\000',
	MicrosoftWordE133RowHeightAtLeast = '\002\025\000\001',
	MicrosoftWordE133RowHeightExactly = '\002\025\000\002'
};
typedef enum MicrosoftWordE133 MicrosoftWordE133;

enum MicrosoftWordE134 {
	MicrosoftWordE134FrameAuto = '\002\026\000\000',
	MicrosoftWordE134FrameAtLeast = '\002\026\000\001',
	MicrosoftWordE134FrameExact = '\002\026\000\002'
};
typedef enum MicrosoftWordE134 MicrosoftWordE134;

enum MicrosoftWordE135 {
	MicrosoftWordE135InsertCellsShiftRight = '\002\027\000\000',
	MicrosoftWordE135InsertCellsShiftDown = '\002\027\000\001',
	MicrosoftWordE135InsertCellsEntireRow = '\002\027\000\002',
	MicrosoftWordE135InsertCellsEntireColumn = '\002\027\000\003'
};
typedef enum MicrosoftWordE135 MicrosoftWordE135;

enum MicrosoftWordE136 {
	MicrosoftWordE136DeleteCellsShiftLeft = '\002\030\000\000',
	MicrosoftWordE136DeleteCellsShiftUp = '\002\030\000\001',
	MicrosoftWordE136DeleteCellsEntireRow = '\002\030\000\002',
	MicrosoftWordE136DeleteCellsEntireColumn = '\002\030\000\003'
};
typedef enum MicrosoftWordE136 MicrosoftWordE136;

enum MicrosoftWordE137 {
	MicrosoftWordE137ListApplyToWholeList = '\002\031\000\000',
	MicrosoftWordE137ListApplyToThisPointForward = '\002\031\000\001',
	MicrosoftWordE137ListApplyToSelection = '\002\031\000\002'
};
typedef enum MicrosoftWordE137 MicrosoftWordE137;

enum MicrosoftWordE138 {
	MicrosoftWordE138AlertsNone = '\002\032\000\000',
	MicrosoftWordE138AlertsMessageBox = '\002\031\377\376',
	MicrosoftWordE138AlertsAll = '\002\031\377\377'
};
typedef enum MicrosoftWordE138 MicrosoftWordE138;

enum MicrosoftWordE139 {
	MicrosoftWordE139CursorWait = '\002\033\000\000',
	MicrosoftWordE139CursorIbeam = '\002\033\000\001',
	MicrosoftWordE139CursorNormal = '\002\033\000\002',
	MicrosoftWordE139CursorNorthwestArrow = '\002\033\000\003'
};
typedef enum MicrosoftWordE139 MicrosoftWordE139;

enum MicrosoftWordE141 {
	MicrosoftWordE141AdjustNone = '\002\035\000\000',
	MicrosoftWordE141AdjustProportional = '\002\035\000\001',
	MicrosoftWordE141AdjustFirstColumn = '\002\035\000\002',
	MicrosoftWordE141AdjustSameWidth = '\002\035\000\003'
};
typedef enum MicrosoftWordE141 MicrosoftWordE141;

enum MicrosoftWordE142 {
	MicrosoftWordE142AlignParagraphLeft = '\002\036\000\000',
	MicrosoftWordE142AlignParagraphCenter = '\002\036\000\001',
	MicrosoftWordE142AlignParagraphRight = '\002\036\000\002',
	MicrosoftWordE142AlignParagraphJustify = '\002\036\000\003',
	MicrosoftWordE142AlignParagraphDistribute = '\002\036\000\004'
};
typedef enum MicrosoftWordE142 MicrosoftWordE142;

enum MicrosoftWordE143 {
	MicrosoftWordE143ListLevelAlignLeft = '\002\037\000\000',
	MicrosoftWordE143ListLevelAlignCenter = '\002\037\000\001',
	MicrosoftWordE143ListLevelAlignRight = '\002\037\000\002'
};
typedef enum MicrosoftWordE143 MicrosoftWordE143;

enum MicrosoftWordE144 {
	MicrosoftWordE144AlignRowLeft = '\002 \000\000',
	MicrosoftWordE144AlignRowCenter = '\002 \000\001',
	MicrosoftWordE144AlignRowRight = '\002 \000\002'
};
typedef enum MicrosoftWordE144 MicrosoftWordE144;

enum MicrosoftWordE145 {
	MicrosoftWordE145AlignTabLeft = '\002!\000\000',
	MicrosoftWordE145AlignTabCenter = '\002!\000\001',
	MicrosoftWordE145AlignTabRight = '\002!\000\002',
	MicrosoftWordE145AlignTabDecimal = '\002!\000\003',
	MicrosoftWordE145AlignTabBar = '\002!\000\004',
	MicrosoftWordE145AlignTabList = '\002!\000\006'
};
typedef enum MicrosoftWordE145 MicrosoftWordE145;

enum MicrosoftWordE146 {
	MicrosoftWordE146AlignVerticalTop = '\002\"\000\000',
	MicrosoftWordE146AlignVerticalCenter = '\002\"\000\001',
	MicrosoftWordE146AlignVerticalJustify = '\002\"\000\002',
	MicrosoftWordE146AlignVerticalBottom = '\002\"\000\003'
};
typedef enum MicrosoftWordE146 MicrosoftWordE146;

enum MicrosoftWordE147 {
	MicrosoftWordE147CellAlignVerticalTop = '\002#\000\000',
	MicrosoftWordE147CellAlignVerticalCenter = '\002#\000\001',
	MicrosoftWordE147CellAlignVerticalBottom = '\002#\000\003'
};
typedef enum MicrosoftWordE147 MicrosoftWordE147;

enum MicrosoftWordE148 {
	MicrosoftWordE148RangeAnnotationAlignmentCenter = '\002$\000\000',
	MicrosoftWordE148ZeroOneZero = '\002$\000\001',
	MicrosoftWordE148OneTwoOne = '\002$\000\002',
	MicrosoftWordE148RangeAnnotationAlignmentLeft = '\002$\000\003',
	MicrosoftWordE148RangeAnnotationAlignmentRight = '\002$\000\004'
};
typedef enum MicrosoftWordE148 MicrosoftWordE148;

enum MicrosoftWordE149 {
	MicrosoftWordE149TrailingTab = '\002%\000\000',
	MicrosoftWordE149TrailingSpace = '\002%\000\001',
	MicrosoftWordE149TrailingNone = '\002%\000\002'
};
typedef enum MicrosoftWordE149 MicrosoftWordE149;

enum MicrosoftWordE150 {
	MicrosoftWordE150BulletGallery = '\002&\000\001',
	MicrosoftWordE150NumberGallery = '\002&\000\002',
	MicrosoftWordE150OutlineNumberGallery = '\002&\000\003'
};
typedef enum MicrosoftWordE150 MicrosoftWordE150;

enum MicrosoftWordE151 {
	MicrosoftWordE151ListNumberStyleArabic = '\002\'\000\000',
	MicrosoftWordE151ListNumberStyleUppercaseRoman = '\002\'\000\001',
	MicrosoftWordE151ListNumberStyleLowercaseRoman = '\002\'\000\002',
	MicrosoftWordE151ListNumberStyleUppercaseLetter = '\002\'\000\003',
	MicrosoftWordE151ListNumberStyleLowercaseLetter = '\002\'\000\004',
	MicrosoftWordE151ListNumberStyleOrdinal = '\002\'\000\005',
	MicrosoftWordE151ListNumberStyleCardinalText = '\002\'\000\006',
	MicrosoftWordE151ListNumberStyleOrdinalText = '\002\'\000\007',
	MicrosoftWordE151ListNumberStyleKanji = '\002\'\000\012',
	MicrosoftWordE151ListNumberStyleKanjiDigit = '\002\'\000\013',
	MicrosoftWordE151ListNumberStyleAiueoHalfWidth = '\002\'\000\014',
	MicrosoftWordE151ListNumberStyleIrohaHalfWidth = '\002\'\000\015',
	MicrosoftWordE151ListNumberStyleArabicFullWidth = '\002\'\000\016',
	MicrosoftWordE151ListNumberStyleKanjiTraditional = '\002\'\000\020',
	MicrosoftWordE151ListNumberStyleKanjiTraditional2 = '\002\'\000\021',
	MicrosoftWordE151ListNumberStyleNumberInCircle = '\002\'\000\022',
	MicrosoftWordE151ListNumberStyleAiueo = '\002\'\000\024',
	MicrosoftWordE151ListNumberStyleIroha = '\002\'\000\025',
	MicrosoftWordE151ListNumberStyleArabicLz = '\002\'\000\026',
	MicrosoftWordE151ListNumberStyleBullet = '\002\'\000\027',
	MicrosoftWordE151ListNumberStyleGanada = '\002\'\000\030',
	MicrosoftWordE151ListNumberStyleChosung = '\002\'\000\031',
	MicrosoftWordE151ListNumberStyleGbnum1 = '\002\'\000\032',
	MicrosoftWordE151ListNumberStyleGbnum2 = '\002\'\000\033',
	MicrosoftWordE151ListNumberStyleGbnum3 = '\002\'\000\034',
	MicrosoftWordE151ListNumberStyleGbnum4 = '\002\'\000\035',
	MicrosoftWordE151ListNumberStyleZodiac1 = '\002\'\000\036',
	MicrosoftWordE151ListNumberStyleZodiac2 = '\002\'\000\037',
	MicrosoftWordE151ListNumberStyleZodiac3 = '\002\'\000 ',
	MicrosoftWordE151ListNumberStyleTradChinNum1 = '\002\'\000!',
	MicrosoftWordE151ListNumberStyleTradChinNum2 = '\002\'\000\"',
	MicrosoftWordE151ListNumberStyleTradChinNum3 = '\002\'\000#',
	MicrosoftWordE151ListNumberStyleTradChinNum4 = '\002\'\000$',
	MicrosoftWordE151ListNumberStyleSimpChinNum1 = '\002\'\000%',
	MicrosoftWordE151ListNumberStyleSimpChinNum2 = '\002\'\000&',
	MicrosoftWordE151ListNumberStyleSimpChinNum3 = '\002\'\000\'',
	MicrosoftWordE151ListNumberStyleSimpChinNum4 = '\002\'\000(',
	MicrosoftWordE151ListNumberStyleHanjaRead = '\002\'\000)',
	MicrosoftWordE151ListNumberStyleHanjaReadDigit = '\002\'\000*',
	MicrosoftWordE151ListNumberStyleHangul = '\002\'\000+',
	MicrosoftWordE151ListNumberStyleHanja = '\002\'\000,',
	MicrosoftWordE151ListNumberStylePictureBullet = '\002\'\000\371',
	MicrosoftWordE151ListNumberStyleLegal = '\002\'\000\375',
	MicrosoftWordE151ListNumberStyleLegalLz = '\002\'\000\376',
	MicrosoftWordE151ListNumberStyleNone = '\002\'\000\377'
};
typedef enum MicrosoftWordE151 MicrosoftWordE151;

enum MicrosoftWordE152 {
	MicrosoftWordE152NoteNumberStyleArabic = '\002(\000\000',
	MicrosoftWordE152NoteNumberStyleUppercaseRoman = '\002(\000\001',
	MicrosoftWordE152NoteNumberStyleLowercaseRoman = '\002(\000\002',
	MicrosoftWordE152NoteNumberStyleUppercaseLetter = '\002(\000\003',
	MicrosoftWordE152NoteNumberStyleLowercaseLetter = '\002(\000\004',
	MicrosoftWordE152NoteNumberStyleSymbol = '\002(\000\011',
	MicrosoftWordE152NoteNumberStyleArabicFullWidth = '\002(\000\016',
	MicrosoftWordE152NoteNumberStyleKanji = '\002(\000\012',
	MicrosoftWordE152NoteNumberStyleKanjiDigit = '\002(\000\013',
	MicrosoftWordE152NoteNumberStyleKanjiTraditional = '\002(\000\020',
	MicrosoftWordE152NoteNumberStyleNumberInCircle = '\002(\000\022',
	MicrosoftWordE152NoteNumberStyleHanjaRead = '\002(\000)',
	MicrosoftWordE152NoteNumberStyleHanjaReadDigit = '\002(\000*',
	MicrosoftWordE152NoteNumberStyleTradChinNum1 = '\002(\000!',
	MicrosoftWordE152NoteNumberStyleTradChinNum2 = '\002(\000\"',
	MicrosoftWordE152NoteNumberStyleSimpChinNum1 = '\002(\000%',
	MicrosoftWordE152NoteNumberStyleSimpChinNum2 = '\002(\000&'
};
typedef enum MicrosoftWordE152 MicrosoftWordE152;

enum MicrosoftWordE153 {
	MicrosoftWordE153CaptionNumberStyleArabic = '\002)\000\000',
	MicrosoftWordE153CaptionNumberStyleUppercaseRoman = '\002)\000\001',
	MicrosoftWordE153CaptionNumberStyleLowercaseRoman = '\002)\000\002',
	MicrosoftWordE153CaptionNumberStyleUppercaseLetter = '\002)\000\003',
	MicrosoftWordE153CaptionNumberStyleLowercaseLetter = '\002)\000\004',
	MicrosoftWordE153CaptionNumberStyleArabicFullWidth = '\002)\000\016',
	MicrosoftWordE153CaptionNumberStyleKanji = '\002)\000\012',
	MicrosoftWordE153CaptionNumberStyleKanjiDigit = '\002)\000\013',
	MicrosoftWordE153CaptionNumberStyleKanjiTraditional = '\002)\000\020',
	MicrosoftWordE153CaptionNumberStyleNumberInCircle = '\002)\000\022',
	MicrosoftWordE153CaptionNumberStyleGanada = '\002)\000\030',
	MicrosoftWordE153CaptionNumberStyleChosung = '\002)\000\031',
	MicrosoftWordE153CaptionNumberStyleZodiac1 = '\002)\000\036',
	MicrosoftWordE153CaptionNumberStyleZodiac2 = '\002)\000\037',
	MicrosoftWordE153CaptionNumberStyleHanjaRead = '\002)\000)',
	MicrosoftWordE153CaptionNumberStyleHanjaReadDigit = '\002)\000*',
	MicrosoftWordE153CaptionNumberStyleTradChinNum2 = '\002)\000\"',
	MicrosoftWordE153CaptionNumberStyleTradChinNum3 = '\002)\000#',
	MicrosoftWordE153CaptionNumberStyleSimpChinNum2 = '\002)\000&',
	MicrosoftWordE153CaptionNumberStyleSimpChinNum3 = '\002)\000\''
};
typedef enum MicrosoftWordE153 MicrosoftWordE153;

enum MicrosoftWordE154 {
	MicrosoftWordE154PageNumberStyleArabic = '\002*\000\000',
	MicrosoftWordE154PageNumberStyleUppercaseRoman = '\002*\000\001',
	MicrosoftWordE154PageNumberStyleLowercaseRoman = '\002*\000\002',
	MicrosoftWordE154PageNumberStyleUppercaseLetter = '\002*\000\003',
	MicrosoftWordE154PageNumberStyleLowercaseLetter = '\002*\000\004',
	MicrosoftWordE154PageNumberStyleArabicFullWidth = '\002*\000\016',
	MicrosoftWordE154PageNumberStyleKanji = '\002*\000\012',
	MicrosoftWordE154PageNumberStyleKanjiDigit = '\002*\000\013',
	MicrosoftWordE154PageNumberStyleKanjiTraditional = '\002*\000\020',
	MicrosoftWordE154PageNumberStyleNumberInCircle = '\002*\000\022',
	MicrosoftWordE154PageNumberStyleHanjaRead = '\002*\000)',
	MicrosoftWordE154PageNumberStyleHanjaReadDigit = '\002*\000*',
	MicrosoftWordE154PageNumberStyleTradChinNum1 = '\002*\000!',
	MicrosoftWordE154PageNumberStyleTradChinNum2 = '\002*\000\"',
	MicrosoftWordE154PageNumberStyleSimpChinNum1 = '\002*\000%',
	MicrosoftWordE154PageNumberStyleSimpChinNum2 = '\002*\000&'
};
typedef enum MicrosoftWordE154 MicrosoftWordE154;

enum MicrosoftWordE155 {
	MicrosoftWordE155StatisticWords = '\002+\000\000',
	MicrosoftWordE155StatisticLines = '\002+\000\001',
	MicrosoftWordE155StatisticPages = '\002+\000\002',
	MicrosoftWordE155StatisticCharacters = '\002+\000\003',
	MicrosoftWordE155StatisticParagraphs = '\002+\000\004',
	MicrosoftWordE155StatisticCharactersWithSpaces = '\002+\000\005',
	MicrosoftWordE155StatisticEastAsianCharacters = '\002+\000\006'
};
typedef enum MicrosoftWordE155 MicrosoftWordE155;

enum MicrosoftWordE156 {
	MicrosoftWordE156PropertyTitle = '\002,\000\001',
	MicrosoftWordE156PropertySubject = '\002,\000\002',
	MicrosoftWordE156PropertyAuthor = '\002,\000\003',
	MicrosoftWordE156PropertyKeywords = '\002,\000\004',
	MicrosoftWordE156PropertyComments = '\002,\000\005',
	MicrosoftWordE156PropertyTemplate = '\002,\000\006',
	MicrosoftWordE156PropertyLastAuthor = '\002,\000\007',
	MicrosoftWordE156PropertyRevision = '\002,\000\010',
	MicrosoftWordE156PropertyAppName = '\002,\000\011',
	MicrosoftWordE156PropertyTimeLastPrinted = '\002,\000\012',
	MicrosoftWordE156PropertyTimeCreated = '\002,\000\013',
	MicrosoftWordE156PropertyTimeLastSaved = '\002,\000\014',
	MicrosoftWordE156PropertyVbatotalEdit = '\002,\000\015',
	MicrosoftWordE156PropertyPages = '\002,\000\016',
	MicrosoftWordE156PropertyWords = '\002,\000\017',
	MicrosoftWordE156PropertyCharacters = '\002,\000\020',
	MicrosoftWordE156PropertySecurity = '\002,\000\021',
	MicrosoftWordE156PropertyCategory = '\002,\000\022',
	MicrosoftWordE156PropertyFormat = '\002,\000\023',
	MicrosoftWordE156PropertyManager = '\002,\000\024',
	MicrosoftWordE156PropertyCompany = '\002,\000\025',
	MicrosoftWordE156PropertyBytes = '\002,\000\026',
	MicrosoftWordE156PropertyLines = '\002,\000\027',
	MicrosoftWordE156PropertyParas = '\002,\000\030',
	MicrosoftWordE156PropertySlides = '\002,\000\031',
	MicrosoftWordE156PropertyNotes = '\002,\000\032',
	MicrosoftWordE156PropertyHiddenSlides = '\002,\000\033',
	MicrosoftWordE156PropertyMmclips = '\002,\000\034',
	MicrosoftWordE156PropertyHyperlinkBase = '\002,\000\035',
	MicrosoftWordE156PropertyCharsWspaces = '\002,\000\036'
};
typedef enum MicrosoftWordE156 MicrosoftWordE156;

enum MicrosoftWordE157 {
	MicrosoftWordE157LineSpaceSingle = '\002-\000\000',
	MicrosoftWordE157LineSpace1Pt5 = '\002-\000\001',
	MicrosoftWordE157LineSpaceDouble = '\002-\000\002',
	MicrosoftWordE157LineSpaceAtLeast = '\002-\000\003',
	MicrosoftWordE157LineSpaceExactly = '\002-\000\004',
	MicrosoftWordE157LineSpaceMultiple = '\002-\000\005'
};
typedef enum MicrosoftWordE157 MicrosoftWordE157;

enum MicrosoftWordE158 {
	MicrosoftWordE158NumberParagraph = '\002.\000\001',
	MicrosoftWordE158NumberListnum = '\002.\000\002',
	MicrosoftWordE158NumberAllNumbers = '\002.\000\003'
};
typedef enum MicrosoftWordE158 MicrosoftWordE158;

enum MicrosoftWordE159 {
	MicrosoftWordE159ListNoNumbering = '\002/\000\000',
	MicrosoftWordE159ListListnumOnly = '\002/\000\001',
	MicrosoftWordE159ListBullet = '\002/\000\002',
	MicrosoftWordE159ListSimpleNumbering = '\002/\000\003',
	MicrosoftWordE159ListOutlineNumbering = '\002/\000\004',
	MicrosoftWordE159ListMixedNumbering = '\002/\000\005',
	MicrosoftWordE159ListPictureBullet = '\002/\000\006'
};
typedef enum MicrosoftWordE159 MicrosoftWordE159;

enum MicrosoftWordE160 {
	MicrosoftWordE160MainTextStory = '\0020\000\001',
	MicrosoftWordE160FootnotesStory = '\0020\000\002',
	MicrosoftWordE160EndnotesStory = '\0020\000\003',
	MicrosoftWordE160CommentsStory = '\0020\000\004',
	MicrosoftWordE160TextFrameStory = '\0020\000\005',
	MicrosoftWordE160EvenPagesHeaderStory = '\0020\000\006',
	MicrosoftWordE160PrimaryHeaderStory = '\0020\000\007',
	MicrosoftWordE160EvenPagesFooterStory = '\0020\000\010',
	MicrosoftWordE160PrimaryFooterStory = '\0020\000\011',
	MicrosoftWordE160FirstPageHeaderStory = '\0020\000\012',
	MicrosoftWordE160FirstPageFooterStory = '\0020\000\013',
	MicrosoftWordE160FootnoteSeparatorStory = '\0020\000\014',
	MicrosoftWordE160FootnoteContinuationSeparatorStory = '\0020\000\015',
	MicrosoftWordE160FootnoteContinuationNoticeStory = '\0020\000\016',
	MicrosoftWordE160EndnoteSeparatorStory = '\0020\000\017',
	MicrosoftWordE160EndnoteContinuationSeparatorStory = '\0020\000\020',
	MicrosoftWordE160EndnoteContinuationNoticeStory = '\0020\000\021'
};
typedef enum MicrosoftWordE160 MicrosoftWordE160;

enum MicrosoftWordE161 {
	MicrosoftWordE161FormatDocument97i = '\0021\000\240',
	MicrosoftWordE161FormatDocument97 = '\0021\000\000',
	MicrosoftWordE161FormatTemplate97i = '\0021\000\241',
	MicrosoftWordE161FormatTemplate97 = '\0021\000\001',
	MicrosoftWordE161FormatText = '\0021\000\002',
	MicrosoftWordE161FormatTextLineBreaks = '\0021\000\003',
	MicrosoftWordE161FormatDostext = '\0021\000\004',
	MicrosoftWordE161FormatDostextLineBreaks = '\0021\000\005',
	MicrosoftWordE161FormatRtf = '\0021\000\006',
	MicrosoftWordE161FormatUnicodeTexti = '\0021\000\247',
	MicrosoftWordE161FormatUnicodeText = '\0021\000\007',
	MicrosoftWordE161FormatHTML = '\0021\000\010',
	MicrosoftWordE161FormatWebArchive = '\0021\000\011',
	MicrosoftWordE161FormatStationery = '\0021\000\012',
	MicrosoftWordE161FormatXml = '\0021\000\013',
	MicrosoftWordE161FormatDocument = '\0021\000\014',
	MicrosoftWordE161FormatDocumentME = '\0021\000\015',
	MicrosoftWordE161FormatTemplate = '\0021\000\016',
	MicrosoftWordE161FormatTemplateME = '\0021\000\017',
	MicrosoftWordE161FormatPDF = '\0021\000\020',
	MicrosoftWordE161FormatFlatDocument = '\0021\000\021',
	MicrosoftWordE161FormatFlatDocumentME = '\0021\000\022',
	MicrosoftWordE161FormatFlatTemplate = '\0021\000\023',
	MicrosoftWordE161FormatFlatTemplateME = '\0021\000\024',
	MicrosoftWordE161FormatCustomDictionary = '\0021\000\025',
	MicrosoftWordE161FormatExcludeDictionary = '\0021\000\026',
	MicrosoftWordE161FormatDocumentAuto = '\0021\015\014',
	MicrosoftWordE161FormatTemplateAuto = '\0021\015\007'
};
typedef enum MicrosoftWordE161 MicrosoftWordE161;

enum MicrosoftWordE162 {
	MicrosoftWordE162OpenFormatAuto = '\0022\000\000',
	MicrosoftWordE162OpenFormatDocument = '\0022\000\001',
	MicrosoftWordE162OpenFormatTemplate = '\0022\000\002',
	MicrosoftWordE162OpenFormatRtf = '\0022\000\003',
	MicrosoftWordE162OpenFormatText = '\0022\000\004',
	MicrosoftWordE162OpenFormatUnicodeText = '\0022\000\005',
	MicrosoftWordE162OpenFormatEncodedText = '\0022\000\005',
	MicrosoftWordE162OpenFormatMacReadable = '\0022\000\006',
	MicrosoftWordE162OpenFormatWebPages = '\0022\000\007',
	MicrosoftWordE162OpenFormatXml = '\0022\000\010',
	MicrosoftWordE162OpenFormatDocument97 = '\0022\000\011',
	MicrosoftWordE162OpenFormatTemplate97 = '\0022\000\012',
	MicrosoftWordE162OpenFormatOffice = '\0022\000\013'
};
typedef enum MicrosoftWordE162 MicrosoftWordE162;

enum MicrosoftWordE163 {
	MicrosoftWordE163HeaderFooterPrimary = '\0023\000\001',
	MicrosoftWordE163HeaderFooterFirstPage = '\0023\000\002',
	MicrosoftWordE163HeaderFooterEvenPages = '\0023\000\003'
};
typedef enum MicrosoftWordE163 MicrosoftWordE163;

enum MicrosoftWordE164 {
	MicrosoftWordE164Toctemplate = '\0024\000\000',
	MicrosoftWordE164Tocclassic = '\0024\000\001',
	MicrosoftWordE164Tocdistinctive = '\0024\000\002',
	MicrosoftWordE164Tocfancy = '\0024\000\003',
	MicrosoftWordE164Tocmodern = '\0024\000\004',
	MicrosoftWordE164Tocformal = '\0024\000\005',
	MicrosoftWordE164Tocsimple = '\0024\000\006'
};
typedef enum MicrosoftWordE164 MicrosoftWordE164;

enum MicrosoftWordE165 {
	MicrosoftWordE165Toftemplate = '\0025\000\000',
	MicrosoftWordE165Tofclassic = '\0025\000\001',
	MicrosoftWordE165Tofdistinctive = '\0025\000\002',
	MicrosoftWordE165Tofcentered = '\0025\000\003',
	MicrosoftWordE165Tofformal = '\0025\000\004',
	MicrosoftWordE165Tofsimple = '\0025\000\005'
};
typedef enum MicrosoftWordE165 MicrosoftWordE165;

enum MicrosoftWordE166 {
	MicrosoftWordE166Toatemplate = '\0026\000\000',
	MicrosoftWordE166Toaclassic = '\0026\000\001',
	MicrosoftWordE166Toadistinctive = '\0026\000\002',
	MicrosoftWordE166Toaformal = '\0026\000\003',
	MicrosoftWordE166Toasimple = '\0026\000\004'
};
typedef enum MicrosoftWordE166 MicrosoftWordE166;

enum MicrosoftWordE167 {
	MicrosoftWordE167LineStyleNone = '\0027\000\000',
	MicrosoftWordE167LineStyleSingle = '\0027\000\001',
	MicrosoftWordE167LineStyleDot = '\0027\000\002',
	MicrosoftWordE167LineStyleDashSmallGap = '\0027\000\003',
	MicrosoftWordE167LineStyleDashLargeGap = '\0027\000\004',
	MicrosoftWordE167LineStyleDashDot = '\0027\000\005',
	MicrosoftWordE167LineStyleDashDotDot = '\0027\000\006',
	MicrosoftWordE167LineStyleDouble = '\0027\000\007',
	MicrosoftWordE167LineStyleTriple = '\0027\000\010',
	MicrosoftWordE167LineStyleThinThickSmallGap = '\0027\000\011',
	MicrosoftWordE167LineStyleThickThinSmallGap = '\0027\000\012',
	MicrosoftWordE167LineStyleThinThickThinSmallGap = '\0027\000\013',
	MicrosoftWordE167LineStyleThinThickMedGap = '\0027\000\014',
	MicrosoftWordE167LineStyleThickThinMedGap = '\0027\000\015',
	MicrosoftWordE167LineStyleThinThickThinMedGap = '\0027\000\016',
	MicrosoftWordE167LineStyleThinThickLargeGap = '\0027\000\017',
	MicrosoftWordE167LineStyleThickThinLargeGap = '\0027\000\020',
	MicrosoftWordE167LineStyleThinThickThinLargeGap = '\0027\000\021',
	MicrosoftWordE167LineStyleSingleWavy = '\0027\000\022',
	MicrosoftWordE167LineStyleDoubleWavy = '\0027\000\023',
	MicrosoftWordE167LineStyleDashDotStroked = '\0027\000\024',
	MicrosoftWordE167LineStyleEmboss_3D = '\0027\000\025',
	MicrosoftWordE167LineStyleEngrave_3D = '\0027\000\026',
	MicrosoftWordE167LineStyleOutset = '\0027\000\027',
	MicrosoftWordE167LineStyleInset = '\0027\000\030'
};
typedef enum MicrosoftWordE167 MicrosoftWordE167;

enum MicrosoftWordE168 {
	MicrosoftWordE168LineWidth25Point = '\0028\000\002',
	MicrosoftWordE168LineWidth50Point = '\0028\000\004',
	MicrosoftWordE168LineWidth75Point = '\0028\000\006',
	MicrosoftWordE168LineWidth100Point = '\0028\000\010',
	MicrosoftWordE168LineWidth150Point = '\0028\000\014',
	MicrosoftWordE168LineWidth225Point = '\0028\000\022',
	MicrosoftWordE168LineWidth300Point = '\0028\000\030',
	MicrosoftWordE168LineWidth450Point = '\0028\000$',
	MicrosoftWordE168LineWidth600Point = '\0028\0000'
};
typedef enum MicrosoftWordE168 MicrosoftWordE168;

enum MicrosoftWordE169 {
	MicrosoftWordE169SectionBreakNextPage = '\0029\000\002',
	MicrosoftWordE169SectionBreakContinuous = '\0029\000\003',
	MicrosoftWordE169SectionBreakEvenPage = '\0029\000\004',
	MicrosoftWordE169SectionBreakOddPage = '\0029\000\005',
	MicrosoftWordE169LineBreak = '\0029\000\006',
	MicrosoftWordE169PageBreak = '\0029\000\007',
	MicrosoftWordE169ColumnBreak = '\0029\000\010'
};
typedef enum MicrosoftWordE169 MicrosoftWordE169;

enum MicrosoftWordE313 {
	MicrosoftWordE313ContinuedMaster = '\002\311\000\000'
};
typedef enum MicrosoftWordE313 MicrosoftWordE313;

enum MicrosoftWordE170 {
	MicrosoftWordE170TabLeaderSpaces = '\002:\000\000',
	MicrosoftWordE170TabLeaderDots = '\002:\000\001',
	MicrosoftWordE170TabLeaderDashes = '\002:\000\002',
	MicrosoftWordE170TabLeaderLines = '\002:\000\003',
	MicrosoftWordE170TabLeaderHeavy = '\002:\000\004',
	MicrosoftWordE170TabLeaderMiddleDot = '\002:\000\005'
};
typedef enum MicrosoftWordE170 MicrosoftWordE170;

enum MicrosoftWordE171 {
	MicrosoftWordE171Inches = '\002;\000\000',
	MicrosoftWordE171Centimeters = '\002;\000\001',
	MicrosoftWordE171Millimeters = '\002;\000\002',
	MicrosoftWordE171Points = '\002;\000\003',
	MicrosoftWordE171Picas = '\002;\000\004'
};
typedef enum MicrosoftWordE171 MicrosoftWordE171;

enum MicrosoftWordE172 {
	MicrosoftWordE172DropNone = '\002<\000\000',
	MicrosoftWordE172DropNormal = '\002<\000\001',
	MicrosoftWordE172DropMargin = '\002<\000\002'
};
typedef enum MicrosoftWordE172 MicrosoftWordE172;

enum MicrosoftWordE173 {
	MicrosoftWordE173RestartContinuous = '\002=\000\000',
	MicrosoftWordE173RestartSection = '\002=\000\001',
	MicrosoftWordE173RestartPage = '\002=\000\002'
};
typedef enum MicrosoftWordE173 MicrosoftWordE173;

enum MicrosoftWordE174 {
	MicrosoftWordE174BottomOfPage = '\002>\000\000',
	MicrosoftWordE174BeneathText = '\002>\000\001'
};
typedef enum MicrosoftWordE174 MicrosoftWordE174;

enum MicrosoftWordE175 {
	MicrosoftWordE175End_of_section = '\002\?\000\000',
	MicrosoftWordE175End_of_document = '\002\?\000\001'
};
typedef enum MicrosoftWordE175 MicrosoftWordE175;

enum MicrosoftWordE176 {
	MicrosoftWordE176SortSeparateByTabs = '\002@\000\000',
	MicrosoftWordE176SortSeparateByCommas = '\002@\000\001',
	MicrosoftWordE176SortSeparateByDefaultTableSeparator = '\002@\000\002'
};
typedef enum MicrosoftWordE176 MicrosoftWordE176;

enum MicrosoftWordE177 {
	MicrosoftWordE177SeparateByParagraphs = '\002A\000\000',
	MicrosoftWordE177SeparateByTabs = '\002A\000\001',
	MicrosoftWordE177SeparateByCommas = '\002A\000\002',
	MicrosoftWordE177SeparateByDefaultListSeparator = '\002A\000\003'
};
typedef enum MicrosoftWordE177 MicrosoftWordE177;

enum MicrosoftWordE178 {
	MicrosoftWordE178SortFieldAlphanumeric = '\002B\000\000',
	MicrosoftWordE178SortFieldNumeric = '\002B\000\001',
	MicrosoftWordE178SortFieldDate = '\002B\000\002',
	MicrosoftWordE178SortFieldSyllable = '\002B\000\003',
	MicrosoftWordE178SortFieldJapanJis = '\002B\000\004',
	MicrosoftWordE178SortFieldStroke = '\002B\000\005',
	MicrosoftWordE178SortFieldKoreaKs = '\002B\000\006'
};
typedef enum MicrosoftWordE178 MicrosoftWordE178;

enum MicrosoftWordE179 {
	MicrosoftWordE179SortOrderAscending = '\002C\000\000',
	MicrosoftWordE179SortOrderDescending = '\002C\000\001'
};
typedef enum MicrosoftWordE179 MicrosoftWordE179;

enum MicrosoftWordE180 {
	MicrosoftWordE180TableFormatNone = '\002D\000\000',
	MicrosoftWordE180TableFormatSimple1 = '\002D\000\001',
	MicrosoftWordE180TableFormatSimple2 = '\002D\000\002',
	MicrosoftWordE180TableFormatSimple3 = '\002D\000\003',
	MicrosoftWordE180TableFormatClassic1 = '\002D\000\004',
	MicrosoftWordE180TableFormatClassic2 = '\002D\000\005',
	MicrosoftWordE180TableFormatClassic3 = '\002D\000\006',
	MicrosoftWordE180TableFormatClassic4 = '\002D\000\007',
	MicrosoftWordE180TableFormatColorful1 = '\002D\000\010',
	MicrosoftWordE180TableFormatColorful2 = '\002D\000\011',
	MicrosoftWordE180TableFormatColorful3 = '\002D\000\012',
	MicrosoftWordE180TableFormatColumns1 = '\002D\000\013',
	MicrosoftWordE180TableFormatColumns2 = '\002D\000\014',
	MicrosoftWordE180TableFormatColumns3 = '\002D\000\015',
	MicrosoftWordE180TableFormatColumns4 = '\002D\000\016',
	MicrosoftWordE180TableFormatColumns5 = '\002D\000\017',
	MicrosoftWordE180TableFormatGrid1 = '\002D\000\020',
	MicrosoftWordE180TableFormatGrid2 = '\002D\000\021',
	MicrosoftWordE180TableFormatGrid3 = '\002D\000\022',
	MicrosoftWordE180TableFormatGrid4 = '\002D\000\023',
	MicrosoftWordE180TableFormatGrid5 = '\002D\000\024',
	MicrosoftWordE180TableFormatGrid6 = '\002D\000\025',
	MicrosoftWordE180TableFormatGrid7 = '\002D\000\026',
	MicrosoftWordE180TableFormatGrid8 = '\002D\000\027',
	MicrosoftWordE180TableFormatList1 = '\002D\000\030',
	MicrosoftWordE180TableFormatList2 = '\002D\000\031',
	MicrosoftWordE180TableFormatList3 = '\002D\000\032',
	MicrosoftWordE180TableFormatList4 = '\002D\000\033',
	MicrosoftWordE180TableFormatList5 = '\002D\000\034',
	MicrosoftWordE180TableFormatList6 = '\002D\000\035',
	MicrosoftWordE180TableFormatList7 = '\002D\000\036',
	MicrosoftWordE180TableFormatList8 = '\002D\000\037',
	MicrosoftWordE180TableFormat3DEffects1 = '\002D\000 ',
	MicrosoftWordE180TableFormat3DEffects2 = '\002D\000!',
	MicrosoftWordE180TableFormat3DEffects3 = '\002D\000\"',
	MicrosoftWordE180TableFormatContemporary = '\002D\000#',
	MicrosoftWordE180TableFormatElegant = '\002D\000$',
	MicrosoftWordE180TableFormatProfessional = '\002D\000%',
	MicrosoftWordE180TableFormatSubtle1 = '\002D\000&',
	MicrosoftWordE180TableFormatSubtle2 = '\002D\000\'',
	MicrosoftWordE180TableFormatWeb1 = '\002D\000(',
	MicrosoftWordE180TableFormatWeb2 = '\002D\000)',
	MicrosoftWordE180TableFormatWeb3 = '\002D\000*'
};
typedef enum MicrosoftWordE180 MicrosoftWordE180;

enum MicrosoftWordE181 {
	MicrosoftWordE181TableFormatApplyBorders = '\002E\000\001',
	MicrosoftWordE181TableFormatApplyShading = '\002E\000\002',
	MicrosoftWordE181TableFormatApplyFont = '\002E\000\004',
	MicrosoftWordE181TableFormatApplyColor = '\002E\000\010',
	MicrosoftWordE181TableFormatApplyAutoFit = '\002E\000\020',
	MicrosoftWordE181TableFormatApplyHeadingRows = '\002E\000 ',
	MicrosoftWordE181TableFormatApplyLastRow = '\002E\000@',
	MicrosoftWordE181TableFormatApplyFirstColumn = '\002E\000\200',
	MicrosoftWordE181TableFormatApplyLastColumn = '\002E\001\000'
};
typedef enum MicrosoftWordE181 MicrosoftWordE181;

enum MicrosoftWordE182 {
	MicrosoftWordE182LanguageNone = '\002F\000\000',
	MicrosoftWordE182LanguageNoProofing = '\002F\004\000',
	MicrosoftWordE182Danish = '\002F\004\006',
	MicrosoftWordE182German = '\002F\004\007',
	MicrosoftWordE182SwissGerman = '\002F\010\007',
	MicrosoftWordE182AustrianGerman = '\002F\014\007',
	MicrosoftWordE182EnglishAus = '\002F\014\011',
	MicrosoftWordE182EnglishUk = '\002F\010\011',
	MicrosoftWordE182EnglishUs = '\002F\004\011',
	MicrosoftWordE182EnglishCanadian = '\002F\020\011',
	MicrosoftWordE182EnglishNewZealand = '\002F\024\011',
	MicrosoftWordE182EnglishSouthAfrica = '\002F\034\011',
	MicrosoftWordE182Spanish = '\002F\004\012',
	MicrosoftWordE182French = '\002F\004\014',
	MicrosoftWordE182FrenchCanadian = '\002F\014\014',
	MicrosoftWordE182Italian = '\002F\004\020',
	MicrosoftWordE182Dutch = '\002F\004\023',
	MicrosoftWordE182NorwegianBokmol = '\002F\004\024',
	MicrosoftWordE182NorwegianNynorsk = '\002F\010\024',
	MicrosoftWordE182BrazilianPortuguese = '\002F\004\026',
	MicrosoftWordE182Portuguese = '\002F\010\026',
	MicrosoftWordE182Finnish = '\002F\004\013',
	MicrosoftWordE182Swedish = '\002F\004\035',
	MicrosoftWordE182Catalan = '\002F\004\003',
	MicrosoftWordE182Greek = '\002F\004\010',
	MicrosoftWordE182Turkish = '\002F\004\037',
	MicrosoftWordE182Russian = '\002F\004\031',
	MicrosoftWordE182Czech = '\002F\004\005',
	MicrosoftWordE182Hungarian = '\002F\004\016',
	MicrosoftWordE182Polish = '\002F\004\025',
	MicrosoftWordE182Slovenian = '\002F\004$',
	MicrosoftWordE182Basque = '\002F\004-',
	MicrosoftWordE182Malaysian = '\002F\004>',
	MicrosoftWordE182Japanese = '\002F\004\021',
	MicrosoftWordE182Korean = '\002F\004\022',
	MicrosoftWordE182SimplifiedChinese = '\002F\010\004',
	MicrosoftWordE182TraditionalChinese = '\002F\004\004',
	MicrosoftWordE182SwissFrench = '\002F\020\014',
	MicrosoftWordE182Sesotho = '\002F\0040',
	MicrosoftWordE182Tsonga = '\002F\0041',
	MicrosoftWordE182Tswana = '\002F\0042',
	MicrosoftWordE182Venda = '\002F\0043',
	MicrosoftWordE182Xhosa = '\002F\0044',
	MicrosoftWordE182Zulu = '\002F\0045',
	MicrosoftWordE182Afrikaans = '\002F\0046',
	MicrosoftWordE182Arabic = '\002F\004\001',
	MicrosoftWordE182Hebrew = '\002F\004\015',
	MicrosoftWordE182Slovak = '\002F\004\033',
	MicrosoftWordE182Farsi = '\002F\004)',
	MicrosoftWordE182Romanian = '\002F\004\030',
	MicrosoftWordE182Croatian = '\002F\004\032',
	MicrosoftWordE182Ukrainian = '\002F\004\"',
	MicrosoftWordE182Byelorussian = '\002F\004#',
	MicrosoftWordE182Estonian = '\002F\004%',
	MicrosoftWordE182Latvian = '\002F\004&',
	MicrosoftWordE182Macedonian = '\002F\004/',
	MicrosoftWordE182SerbianLatin = '\002F\010\032',
	MicrosoftWordE182SerbianCyrillic = '\002F\014\032',
	MicrosoftWordE182Icelandic = '\002F\004\017',
	MicrosoftWordE182BelgianFrench = '\002F\010\014',
	MicrosoftWordE182BelgianDutch = '\002F\010\023',
	MicrosoftWordE182Bulgarian = '\002F\004\002',
	MicrosoftWordE182MexicanSpanish = '\002F\010\012',
	MicrosoftWordE182SpanishModernSort = '\002F\014\012',
	MicrosoftWordE182SwissItalian = '\002F\010\020'
};
typedef enum MicrosoftWordE182 MicrosoftWordE182;

enum MicrosoftWordE183 {
	MicrosoftWordE183FieldEmpty = '\002F\377\377',
	MicrosoftWordE183FieldRef = '\002G\000\003',
	MicrosoftWordE183FieldIndexEntry = '\002G\000\004',
	MicrosoftWordE183FieldFootnoteRef = '\002G\000\005',
	MicrosoftWordE183FieldSet = '\002G\000\006',
	MicrosoftWordE183FieldIf = '\002G\000\007',
	MicrosoftWordE183FieldIndex = '\002G\000\010',
	MicrosoftWordE183FieldTocEntry = '\002G\000\011',
	MicrosoftWordE183FieldStyleRef = '\002G\000\012',
	MicrosoftWordE183FieldRefDoc = '\002G\000\013',
	MicrosoftWordE183FieldSequence = '\002G\000\014',
	MicrosoftWordE183FieldToc = '\002G\000\015',
	MicrosoftWordE183FieldInfo = '\002G\000\016',
	MicrosoftWordE183FieldTitle = '\002G\000\017',
	MicrosoftWordE183FieldSubject = '\002G\000\020',
	MicrosoftWordE183FieldAuthor = '\002G\000\021',
	MicrosoftWordE183FieldKeyWord = '\002G\000\022',
	MicrosoftWordE183FieldComments = '\002G\000\023',
	MicrosoftWordE183FieldLastSavedBy = '\002G\000\024',
	MicrosoftWordE183FieldCreateDate = '\002G\000\025',
	MicrosoftWordE183FieldSaveDate = '\002G\000\026',
	MicrosoftWordE183FieldPrintDate = '\002G\000\027',
	MicrosoftWordE183FieldRevisionNum = '\002G\000\030',
	MicrosoftWordE183FieldEditTime = '\002G\000\031',
	MicrosoftWordE183FieldNumPages = '\002G\000\032',
	MicrosoftWordE183FieldNumWords = '\002G\000\033',
	MicrosoftWordE183FieldNumChars = '\002G\000\034',
	MicrosoftWordE183FieldFileName = '\002G\000\035',
	MicrosoftWordE183FieldTemplate = '\002G\000\036',
	MicrosoftWordE183FieldDate = '\002G\000\037',
	MicrosoftWordE183FieldTime = '\002G\000 ',
	MicrosoftWordE183FieldPage = '\002G\000!',
	MicrosoftWordE183FieldExpression = '\002G\000\"',
	MicrosoftWordE183FieldQuote = '\002G\000#',
	MicrosoftWordE183FieldInclude = '\002G\000$',
	MicrosoftWordE183FieldPageRef = '\002G\000%',
	MicrosoftWordE183FieldAsk = '\002G\000&',
	MicrosoftWordE183FieldFillIn = '\002G\000\'',
	MicrosoftWordE183FieldData = '\002G\000(',
	MicrosoftWordE183FieldNext = '\002G\000)',
	MicrosoftWordE183FieldNextIf = '\002G\000*',
	MicrosoftWordE183FieldSkipIf = '\002G\000+',
	MicrosoftWordE183FieldMergeRec = '\002G\000,',
	MicrosoftWordE183FieldDde = '\002G\000-',
	MicrosoftWordE183FieldDdeauto = '\002G\000.',
	MicrosoftWordE183FieldGlossary = '\002G\000/',
	MicrosoftWordE183FieldPrint = '\002G\0000',
	MicrosoftWordE183FieldFormula = '\002G\0001',
	MicrosoftWordE183FieldGoToButton = '\002G\0002',
	MicrosoftWordE183Def0 = '\002G\0003',
	MicrosoftWordE183FieldAutoNumOutline = '\002G\0004',
	MicrosoftWordE183FieldAutoNumLegal = '\002G\0005',
	MicrosoftWordE183FieldAutoNum = '\002G\0006',
	MicrosoftWordE183FieldImport = '\002G\0007',
	MicrosoftWordE183FieldLink = '\002G\0008',
	MicrosoftWordE183FieldSymbol = '\002G\0009',
	MicrosoftWordE183FieldEmbed = '\002G\000:',
	MicrosoftWordE183FieldMergeField = '\002G\000;',
	MicrosoftWordE183FieldUserName = '\002G\000<',
	MicrosoftWordE183FieldUserInitials = '\002G\000=',
	MicrosoftWordE183FieldUserAddress = '\002G\000>',
	MicrosoftWordE183FieldBarCode = '\002G\000\?',
	MicrosoftWordE183FieldDocVariable = '\002G\000@',
	MicrosoftWordE183FieldSection = '\002G\000A',
	MicrosoftWordE183FieldSectionPages = '\002G\000B',
	MicrosoftWordE183FieldIncludePicture = '\002G\000C',
	MicrosoftWordE183FieldIncludeText = '\002G\000D',
	MicrosoftWordE183FieldFileSize = '\002G\000E',
	MicrosoftWordE183FieldFormTextInput = '\002G\000F',
	MicrosoftWordE183FieldFormCheckBox = '\002G\000G',
	MicrosoftWordE183FieldNoteRef = '\002G\000H',
	MicrosoftWordE183FieldToa = '\002G\000I',
	MicrosoftWordE183FieldToaentry = '\002G\000J',
	MicrosoftWordE183FieldMergeSeq = '\002G\000K',
	MicrosoftWordE183FieldPrivate = '\002G\000M',
	MicrosoftWordE183FieldDatabase = '\002G\000N',
	MicrosoftWordE183FieldAutoText = '\002G\000O',
	MicrosoftWordE183FieldCompare = '\002G\000P',
	MicrosoftWordE183FieldAddin = '\002G\000Q',
	MicrosoftWordE183FieldSubscriber = '\002G\000R',
	MicrosoftWordE183FieldFormDropDown = '\002G\000S',
	MicrosoftWordE183FieldAdvance = '\002G\000T',
	MicrosoftWordE183FieldDocProperty = '\002G\000U',
	MicrosoftWordE183FieldOcx = '\002G\000W',
	MicrosoftWordE183FieldHyperlink = '\002G\000X',
	MicrosoftWordE183FieldAutoTextList = '\002G\000Y',
	MicrosoftWordE183FieldListnum = '\002G\000Z',
	MicrosoftWordE183FieldHtmlactiveX = '\002G\000[',
	MicrosoftWordE183FieldContact = '\002G\000b',
	MicrosoftWordE183FieldUserProperty = '\002G\000c'
};
typedef enum MicrosoftWordE183 MicrosoftWordE183;

enum MicrosoftWordE184 {
	MicrosoftWordE184StyleNormal = '\002G\377\377',
	MicrosoftWordE184StyleEnvelopeAddress = '\002G\377\333',
	MicrosoftWordE184StyleEnvelopeReturn = '\002G\377\332',
	MicrosoftWordE184StyleBodyText = '\002G\377\275',
	MicrosoftWordE184StyleHeading1 = '\002G\377\376',
	MicrosoftWordE184StyleHeading2 = '\002G\377\375',
	MicrosoftWordE184StyleHeading3 = '\002G\377\374',
	MicrosoftWordE184StyleHeading4 = '\002G\377\373',
	MicrosoftWordE184StyleHeading5 = '\002G\377\372',
	MicrosoftWordE184StyleHeading6 = '\002G\377\371',
	MicrosoftWordE184StyleHeading7 = '\002G\377\370',
	MicrosoftWordE184StyleHeading8 = '\002G\377\367',
	MicrosoftWordE184StyleHeading9 = '\002G\377\366',
	MicrosoftWordE184StyleIndex1 = '\002G\377\365',
	MicrosoftWordE184StyleIndex2 = '\002G\377\364',
	MicrosoftWordE184StyleIndex3 = '\002G\377\363',
	MicrosoftWordE184StyleIndex4 = '\002G\377\362',
	MicrosoftWordE184StyleIndex5 = '\002G\377\361',
	MicrosoftWordE184StyleIndex6 = '\002G\377\360',
	MicrosoftWordE184StyleIndex7 = '\002G\377\357',
	MicrosoftWordE184StyleIndex8 = '\002G\377\356',
	MicrosoftWordE184StyleIndex9 = '\002G\377\355',
	MicrosoftWordE184StyleToc1 = '\002G\377\354',
	MicrosoftWordE184StyleToc2 = '\002G\377\353',
	MicrosoftWordE184StyleToc3 = '\002G\377\352',
	MicrosoftWordE184StyleToc4 = '\002G\377\351',
	MicrosoftWordE184StyleToc5 = '\002G\377\350',
	MicrosoftWordE184StyleToc6 = '\002G\377\347',
	MicrosoftWordE184StyleToc7 = '\002G\377\346',
	MicrosoftWordE184StyleToc8 = '\002G\377\345',
	MicrosoftWordE184StyleToc9 = '\002G\377\344',
	MicrosoftWordE184StyleNormalIndent = '\002G\377\343',
	MicrosoftWordE184StyleFootnoteText = '\002G\377\342',
	MicrosoftWordE184StyleCommentText = '\002G\377\341',
	MicrosoftWordE184StyleHeader = '\002G\377\340',
	MicrosoftWordE184StyleFooter = '\002G\377\337',
	MicrosoftWordE184StyleIndexHeading = '\002G\377\336',
	MicrosoftWordE184StyleCaption = '\002G\377\335',
	MicrosoftWordE184StyleTableOfFigures = '\002G\377\334',
	MicrosoftWordE184StyleFootnoteReference = '\002G\377\331',
	MicrosoftWordE184StyleCommentReference = '\002G\377\330',
	MicrosoftWordE184StyleLineNumber = '\002G\377\327',
	MicrosoftWordE184StylePageNumber = '\002G\377\326',
	MicrosoftWordE184StyleEndnoteReference = '\002G\377\325',
	MicrosoftWordE184StyleEndnoteText = '\002G\377\324',
	MicrosoftWordE184StyleTableOfAuthorities = '\002G\377\323',
	MicrosoftWordE184StyleMacroText = '\002G\377\322',
	MicrosoftWordE184StyleToaHeading = '\002G\377\321',
	MicrosoftWordE184StyleList = '\002G\377\320',
	MicrosoftWordE184StyleListBullet = '\002G\377\317',
	MicrosoftWordE184StyleListNumber = '\002G\377\316',
	MicrosoftWordE184StyleList2 = '\002G\377\315',
	MicrosoftWordE184StyleList3 = '\002G\377\314',
	MicrosoftWordE184StyleList4 = '\002G\377\313',
	MicrosoftWordE184StyleList5 = '\002G\377\312',
	MicrosoftWordE184StyleListBullet2 = '\002G\377\311',
	MicrosoftWordE184StyleListBullet3 = '\002G\377\310',
	MicrosoftWordE184StyleListBullet4 = '\002G\377\307',
	MicrosoftWordE184StyleListBullet5 = '\002G\377\306',
	MicrosoftWordE184StyleListNumber2 = '\002G\377\305',
	MicrosoftWordE184StyleListNumber3 = '\002G\377\304',
	MicrosoftWordE184StyleListNumber4 = '\002G\377\303',
	MicrosoftWordE184StyleListNumber5 = '\002G\377\302',
	MicrosoftWordE184StyleTitle = '\002G\377\301',
	MicrosoftWordE184StyleClosing = '\002G\377\300',
	MicrosoftWordE184StyleSignature = '\002G\377\277',
	MicrosoftWordE184StyleDefaultParagraphFont = '\002G\377\276',
	MicrosoftWordE184StyleBodyTextIndent = '\002G\377\274',
	MicrosoftWordE184StyleListContinue = '\002G\377\273',
	MicrosoftWordE184StyleListContinue2 = '\002G\377\272',
	MicrosoftWordE184StyleListContinue3 = '\002G\377\271',
	MicrosoftWordE184StyleListContinue4 = '\002G\377\270',
	MicrosoftWordE184StyleListContinue5 = '\002G\377\267',
	MicrosoftWordE184StyleMessageHeader = '\002G\377\266',
	MicrosoftWordE184StyleSubtitle = '\002G\377\265',
	MicrosoftWordE184StyleSalutation = '\002G\377\264',
	MicrosoftWordE184StyleDate = '\002G\377\263',
	MicrosoftWordE184StyleBodyTextFirstIndent = '\002G\377\262',
	MicrosoftWordE184StyleBodyTextFirstIndent2 = '\002G\377\261',
	MicrosoftWordE184StyleNoteHeading = '\002G\377\260',
	MicrosoftWordE184StyleBodyText2 = '\002G\377\257',
	MicrosoftWordE184StyleBodyText3 = '\002G\377\256',
	MicrosoftWordE184StyleBodyTextIndent2 = '\002G\377\255',
	MicrosoftWordE184StyleBodyTextIndent3 = '\002G\377\254',
	MicrosoftWordE184StyleBlockQuotation = '\002G\377\253',
	MicrosoftWordE184StyleHyperlink = '\002G\377\252',
	MicrosoftWordE184StyleHyperlinkFollowed = '\002G\377\251',
	MicrosoftWordE184StyleStrong = '\002G\377\250',
	MicrosoftWordE184StyleEmphasis = '\002G\377\247',
	MicrosoftWordE184StyleNavPane = '\002G\377\246',
	MicrosoftWordE184StylePlainText = '\002G\377\245',
	MicrosoftWordE184StyleHtmlNormal = '\002G\377\241',
	MicrosoftWordE184StyleHtmlAcronym = '\002G\377\240',
	MicrosoftWordE184StyleHtmlAddress = '\002G\377\237',
	MicrosoftWordE184StyleHtmlCite = '\002G\377\236',
	MicrosoftWordE184StyleHtmlCode = '\002G\377\235',
	MicrosoftWordE184StyleHtmlDefine = '\002G\377\234',
	MicrosoftWordE184StyleHtmlKeyboard = '\002G\377\233',
	MicrosoftWordE184StyleHtmlPreformatted = '\002G\377\232',
	MicrosoftWordE184StyleHtmlSample = '\002G\377\231',
	MicrosoftWordE184StyleHtmlTypewriter = '\002G\377\230',
	MicrosoftWordE184StyleHtmlVariable = '\002G\377\227',
	MicrosoftWordE184StyleNoteLevel1 = '\002G\377c',
	MicrosoftWordE184StyleNoteLevel2 = '\002G\377b',
	MicrosoftWordE184StyleNoteLevel3 = '\002G\377a',
	MicrosoftWordE184StyleNoteLevel4 = '\002G\377`',
	MicrosoftWordE184StyleNoteLevel5 = '\002G\377_',
	MicrosoftWordE184StyleNoteLevel6 = '\002G\377^',
	MicrosoftWordE184StyleNoteLevel7 = '\002G\377]',
	MicrosoftWordE184StyleNoteLevel8 = '\002G\377\\',
	MicrosoftWordE184StyleNoteLevel9 = '\002G\377[',
	MicrosoftWordE184StyleBibliography = '\002G\377Z',
	MicrosoftWordE184StyleListParagraph = '\002G\377Y',
	MicrosoftWordE184StylePlaceholderText = '\002G\377X'
};
typedef enum MicrosoftWordE184 MicrosoftWordE184;

enum MicrosoftWordE185 {
	MicrosoftWordE185DialogFileDocumentLayoutTabMargins = '\002I\000\017',
	MicrosoftWordE185DialogFileDocumentLayoutTabLayout = '\002I\000\020',
	MicrosoftWordE185DialogFilePageSetupTabMargins = '\002I\000\021',
	MicrosoftWordE185DialogFilePageSetupTabPaperSize = '\002I\000\022',
	MicrosoftWordE185DialogFilePageSetupTabPaperSource = '\002I\000\023',
	MicrosoftWordE185DialogFilePageSetupTabLayout = '\002I\000\024',
	MicrosoftWordE185DialogFilePageSetupTabCharsLines = '\002I\000\025',
	MicrosoftWordE185DialogInsertSymbolTabSymbols = '\002I\000\026',
	MicrosoftWordE185DialogInsertSymbolTabSpecialCharacters = '\002I\000\027',
	MicrosoftWordE185DialogNoteOptionsTabAllFootnotes = '\002I\000\030',
	MicrosoftWordE185DialogNoteOptionsTabAllEndnotes = '\002I\000\031',
	MicrosoftWordE185DialogInsertIndexAndTablesTabIndex = '\002I\000\032',
	MicrosoftWordE185DialogInsertIndexAndTablesTabTableOfContents = '\002I\000\033',
	MicrosoftWordE185DialogInsertIndexAndTablesTabTableOfFigures = '\002I\000\034',
	MicrosoftWordE185DialogInsertIndexAndTablesTabTableOfAuthorities = '\002I\000\035',
	MicrosoftWordE185DialogOrganizerTabStyles = '\002I\000\036',
	MicrosoftWordE185DialogOrganizerTabAutoText = '\002I\000\037',
	MicrosoftWordE185DialogOrganizerTabCommandBars = '\002I\000 ',
	MicrosoftWordE185DialogOrganizerTabMacros = '\002I\000!',
	MicrosoftWordE185DialogFormatFontTabFont = '\002I\000\"',
	MicrosoftWordE185DialogFormatFontTabCharacterSpacing = '\002I\000#',
	MicrosoftWordE185DialogFormatBordersAndShadingTabBorders = '\002I\000%',
	MicrosoftWordE185DialogFormatBordersAndShadingTabPageBorder = '\002I\000&',
	MicrosoftWordE185DialogFormatBordersAndShadingTabShading = '\002I\000\'',
	MicrosoftWordE185DialogToolsEnvelopesAndLabelsTabEnvelopes = '\002I\000(',
	MicrosoftWordE185DialogToolsEnvelopesAndLabelsTabLabels = '\002I\000)',
	MicrosoftWordE185DialogFormatParagraphTabIndentsAndSpacing = '\002I\000*',
	MicrosoftWordE185DialogFormatParagraphTabTextFlow = '\002I\000+',
	MicrosoftWordE185DialogFormatParagraphTabTeisai = '\002I\000,',
	MicrosoftWordE185DialogFormatDrawingObjectTabColorsAndLines = '\002I\000-',
	MicrosoftWordE185DialogFormatDrawingObjectTabSize = '\002I\000.',
	MicrosoftWordE185DialogFormatDrawingObjectTabPosition = '\002I\000/',
	MicrosoftWordE185DialogFormatDrawingObjectTabWrapping = '\002I\0000',
	MicrosoftWordE185DialogFormatDrawingObjectTabPicture = '\002I\0001',
	MicrosoftWordE185DialogFormatDrawingObjectTabTextbox = '\002I\0002',
	MicrosoftWordE185DialogFormatDrawingObjectTabHr = '\002I\0003',
	MicrosoftWordE185DialogToolsAutocorrectExceptionsTabFirstLetter = '\002I\0004',
	MicrosoftWordE185DialogToolsAutocorrectExceptionsTabInitialCaps = '\002I\0005',
	MicrosoftWordE185DialogToolsAutocorrectExceptionsTabHangulAndAlphabet = '\002I\0006',
	MicrosoftWordE185DialogToolsAutocorrectExceptionsTabIac = '\002I\0007',
	MicrosoftWordE185DialogFormatBulletsAndNumberingTabBulleted = '\002I\0008',
	MicrosoftWordE185DialogFormatBulletsAndNumberingTabNumbered = '\002I\0009',
	MicrosoftWordE185DialogFormatBulletsAndNumberingTabOutlineNumbered = '\002I\000:',
	MicrosoftWordE185DialogLetterWizardTabLetterFormat = '\002I\000;',
	MicrosoftWordE185DialogLetterWizardTabRecipientInfo = '\002I\000<',
	MicrosoftWordE185DialogLetterWizardTabOtherElements = '\002I\000=',
	MicrosoftWordE185DialogLetterWizardTabSenderInfo = '\002I\000>',
	MicrosoftWordE185DialogToolsAutoManagerTabAutocorrect = '\002I\000\?',
	MicrosoftWordE185DialogToolsAutoManagerTabMathAutocorrect = '\002I\000@',
	MicrosoftWordE185DialogToolsAutoManagerTabAutoFormatAsYouType = '\002I\000A',
	MicrosoftWordE185DialogToolsAutoManagerTabAutoText = '\002I\000B',
	MicrosoftWordE185DialogToolsAutoManagerTabAutoFormat = '\002I\000C',
	MicrosoftWordE185DialogWebOptionsGeneral = '\002I\000D',
	MicrosoftWordE185DialogWebOptionsFiles = '\002I\000E',
	MicrosoftWordE185DialogWebOptionsPictures = '\002I\000F',
	MicrosoftWordE185DialogWebOptionsEncoding = '\002I\000G',
	MicrosoftWordE185DialogWebOptionsFonts = '\002I\000H'
};
typedef enum MicrosoftWordE185 MicrosoftWordE185;

enum MicrosoftWordE186 {
	MicrosoftWordE186DialogHelpAbout = '\002J\000\011',
	MicrosoftWordE186DialogHelpWordPerfectHelp = '\002J\000\012',
	MicrosoftWordE186DialogHelpWordPerfectHelpOptions = '\002J\001\377',
	MicrosoftWordE186DialogFormatChangeCase = '\002J\001B',
	MicrosoftWordE186DialogToolsOptionsFuzzy = '\002J\003\026',
	MicrosoftWordE186DialogToolsWordCount = '\002J\000\344',
	MicrosoftWordE186DialogDocumentStatistics = '\002J\000N',
	MicrosoftWordE186DialogFileNew = '\002J\000O',
	MicrosoftWordE186DialogFileOpen = '\002J\000P',
	MicrosoftWordE186DialogDataMergeOpenDataSource = '\002J\000Q',
	MicrosoftWordE186DialogDataMergeOpenHeaderSource = '\002J\000R',
	MicrosoftWordE186DialogDataMergeUseAddressBook = '\002J\003\013',
	MicrosoftWordE186DialogFileSaveAs = '\002J\000T',
	MicrosoftWordE186DialogFileSummaryInfo = '\002J\000V',
	MicrosoftWordE186DialogToolsTemplates = '\002J\000W',
	MicrosoftWordE186DialogOrganizer = '\002J\000\336',
	MicrosoftWordE186DialogFilePrint = '\002J\000X',
	MicrosoftWordE186DialogDataMerge = '\002J\002\244',
	MicrosoftWordE186DialogDataMergeCheck = '\002J\002\245',
	MicrosoftWordE186DialogDataMergeQueryOptions = '\002J\002\251',
	MicrosoftWordE186DialogDataMergeFindRecord = '\002J\0029',
	MicrosoftWordE186DialogDataMergeInsertIf = '\002J\017\321',
	MicrosoftWordE186DialogDataMergeInsertNextIf = '\002J\017\325',
	MicrosoftWordE186DialogDataMergeInsertSkipIf = '\002J\017\327',
	MicrosoftWordE186DialogDataMergeInsertFillIn = '\002J\017\320',
	MicrosoftWordE186DialogDataMergeInsertAsk = '\002J\017\317',
	MicrosoftWordE186DialogDataMergeInsertSet = '\002J\017\326',
	MicrosoftWordE186DialogDataMergeHelper = '\002J\002\250',
	MicrosoftWordE186DialogLetterWizard = '\002J\0035',
	MicrosoftWordE186DialogFilePrintSetup = '\002J\000a',
	MicrosoftWordE186DialogFileFind = '\002J\000c',
	MicrosoftWordE186DialogDataMergeCreateDataSource = '\002J\002\202',
	MicrosoftWordE186DialogDataMergeCreateHeaderSource = '\002J\002\203',
	MicrosoftWordE186DialogEditPasteSpecial = '\002J\000o',
	MicrosoftWordE186DialogEditFind = '\002J\000p',
	MicrosoftWordE186DialogEditReplace = '\002J\000u',
	MicrosoftWordE186DialogEditGoToOld = '\002J\003+',
	MicrosoftWordE186DialogEditGoTo = '\002J\003\200',
	MicrosoftWordE186DialogCreateAutoText = '\002J\003h',
	MicrosoftWordE186DialogEditAutoText = '\002J\003\331',
	MicrosoftWordE186DialogEditLinks = '\002J\000|',
	MicrosoftWordE186DialogEditObject = '\002J\000}',
	MicrosoftWordE186DialogConvertObject = '\002J\001\210',
	MicrosoftWordE186DialogTableToText = '\002J\000\200',
	MicrosoftWordE186DialogTextToTable = '\002J\000\177',
	MicrosoftWordE186DialogTableInsertTable = '\002J\000\201',
	MicrosoftWordE186DialogTableInsertCells = '\002J\000\202',
	MicrosoftWordE186DialogTableInsertRow = '\002J\000\203',
	MicrosoftWordE186DialogTableDeleteCells = '\002J\000\205',
	MicrosoftWordE186DialogTableSplitCells = '\002J\000\211',
	MicrosoftWordE186DialogTableFormula = '\002J\001\\',
	MicrosoftWordE186DialogTableAutoFormat = '\002J\0023',
	MicrosoftWordE186DialogTableFormatCell = '\002J\002d',
	MicrosoftWordE186DialogViewZoom = '\002J\002A',
	MicrosoftWordE186DialogNewToolbar = '\002J\002J',
	MicrosoftWordE186DialogInsertBreak = '\002J\000\237',
	MicrosoftWordE186DialogInsertFootnote = '\002J\001r',
	MicrosoftWordE186DialogInsertSymbol = '\002J\000\242',
	MicrosoftWordE186DialogInsertPicture = '\002J\000\243',
	MicrosoftWordE186DialogInsertFile = '\002J\000\244',
	MicrosoftWordE186DialogInsertDateTime = '\002J\000\245',
	MicrosoftWordE186DialogInsertNumber = '\002J\003,',
	MicrosoftWordE186DialogInsertField = '\002J\000\246',
	MicrosoftWordE186DialogInsertDatabase = '\002J\001U',
	MicrosoftWordE186DialogInsertMergeField = '\002J\000\247',
	MicrosoftWordE186DialogInsertBookmark = '\002J\000\250',
	MicrosoftWordE186DialogMarkIndexEntry = '\002J\000\251',
	MicrosoftWordE186DialogMarkCitation = '\002J\001\317',
	MicrosoftWordE186DialogEditToacategory = '\002J\002q',
	MicrosoftWordE186DialogInsertIndexAndTables = '\002J\001\331',
	MicrosoftWordE186DialogInsertIndex = '\002J\000\252',
	MicrosoftWordE186DialogInsertTableOfContents = '\002J\000\253',
	MicrosoftWordE186DialogMarkTableOfContentsEntry = '\002J\001\272',
	MicrosoftWordE186DialogInsertTableOfFigures = '\002J\001\330',
	MicrosoftWordE186DialogInsertTableOfAuthorities = '\002J\001\327',
	MicrosoftWordE186DialogInsertObject = '\002J\000\254',
	MicrosoftWordE186DialogFormatCallout = '\002J\002b',
	MicrosoftWordE186DialogDrawSnapToGrid = '\002J\002y',
	MicrosoftWordE186DialogDrawAlign = '\002J\002z',
	MicrosoftWordE186DialogToolsEnvelopesAndLabels = '\002J\002_',
	MicrosoftWordE186DialogToolsCreateEnvelope = '\002J\000\255',
	MicrosoftWordE186DialogToolsCreateLabels = '\002J\001\351',
	MicrosoftWordE186DialogToolsProtectDocument = '\002J\001\367',
	MicrosoftWordE186DialogToolsProtectSection = '\002J\002B',
	MicrosoftWordE186DialogToolsUnprotectDocument = '\002J\002\011',
	MicrosoftWordE186DialogFormatFont = '\002J\000\256',
	MicrosoftWordE186DialogFormatParagraph = '\002J\000\257',
	MicrosoftWordE186DialogFormatSectionLayout = '\002J\000\260',
	MicrosoftWordE186DialogFormatColumns = '\002J\000\261',
	MicrosoftWordE186DialogFileDocumentLayout = '\002J\000\262',
	MicrosoftWordE186DialogFileMacPageSetup = '\002J\002\255',
	MicrosoftWordE186DialogFilePageSetup = '\002J\000\262',
	MicrosoftWordE186DialogFormatTabs = '\002J\000\263',
	MicrosoftWordE186DialogFormatStyle = '\002J\000\264',
	MicrosoftWordE186DialogFormatStyleGallery = '\002J\001\371',
	MicrosoftWordE186DialogFormatDefineStyleFont = '\002J\000\265',
	MicrosoftWordE186DialogFormatDefineStylePara = '\002J\000\266',
	MicrosoftWordE186DialogFormatDefineStyleTabs = '\002J\000\267',
	MicrosoftWordE186DialogFormatDefineStyleFrame = '\002J\000\270',
	MicrosoftWordE186DialogFormatDefineStyleBorders = '\002J\000\271',
	MicrosoftWordE186DialogFormatDefineStyleLang = '\002J\000\272',
	MicrosoftWordE186DialogFormatPicture = '\002J\000\273',
	MicrosoftWordE186DialogToolsLanguage = '\002J\000\274',
	MicrosoftWordE186DialogFormatBordersAndShading = '\002J\000\275',
	MicrosoftWordE186DialogFormatDrawingObject = '\002J\003\300',
	MicrosoftWordE186DialogFormatFrame = '\002J\000\276',
	MicrosoftWordE186DialogFormatDropCap = '\002J\001\350',
	MicrosoftWordE186DialogFormatBulletsAndNumbering = '\002J\0038',
	MicrosoftWordE186DialogToolsHyphenation = '\002J\000\303',
	MicrosoftWordE186DialogToolsBulletsNumbers = '\002J\000\304',
	MicrosoftWordE186DialogToolsHighlightChanges = '\002J\000\305',
	MicrosoftWordE186DialogToolsAcceptRejectChanges = '\002J\001\372',
	MicrosoftWordE186DialogToolsMergeDocuments = '\002J\001\263',
	MicrosoftWordE186DialogToolsCompareDocuments = '\002J\000\306',
	MicrosoftWordE186DialogTableSort = '\002J\000\307',
	MicrosoftWordE186DialogToolsCustomizeMenuBar = '\002J\002g',
	MicrosoftWordE186DialogToolsCustomize = '\002J\000\230',
	MicrosoftWordE186DialogToolsCustomizeKeyboard = '\002J\001\260',
	MicrosoftWordE186DialogToolsCustomizeMenus = '\002J\001\261',
	MicrosoftWordE186DialogListCommands = '\002J\002\323',
	MicrosoftWordE186DialogToolsOptions = '\002J\003\316',
	MicrosoftWordE186DialogToolsOptionsGeneral = '\002J\000\313',
	MicrosoftWordE186DialogToolsAdvancedSettings = '\002J\000\316',
	MicrosoftWordE186DialogToolsOptionsCompatibility = '\002J\002\015',
	MicrosoftWordE186DialogToolsOptionsPrint = '\002J\000\320',
	MicrosoftWordE186DialogToolsOptionsSave = '\002J\000\321',
	MicrosoftWordE186DialogToolsOptionsSpellingAndGrammar = '\002J\000\323',
	MicrosoftWordE186DialogToolsSpellingAndGrammar = '\002J\003<',
	MicrosoftWordE186DialogToolsThesaurus = '\002J\000\302',
	MicrosoftWordE186DialogToolsOptionsUserInfo = '\002J\000\325',
	MicrosoftWordE186DialogToolsOptionsAutoFormat = '\002J\003\277',
	MicrosoftWordE186DialogToolsOptionsTrackChanges = '\002J\001\202',
	MicrosoftWordE186DialogToolsOptionsEdit = '\002J\000\340',
	MicrosoftWordE186DialogInsertPageNumbers = '\002J\001&',
	MicrosoftWordE186DialogFormatPageNumber = '\002J\001*',
	MicrosoftWordE186DialogNoteOptions = '\002J\001u',
	MicrosoftWordE186DialogCopyFile = '\002J\001,',
	MicrosoftWordE186DialogFormatAddrFonts = '\002J\000g',
	MicrosoftWordE186DialogFormatRetAddrFonts = '\002J\000\335',
	MicrosoftWordE186DialogToolsOptionsFileLocations = '\002J\000\341',
	MicrosoftWordE186DialogToolsCreateDirectory = '\002J\003A',
	MicrosoftWordE186DialogUpdateToc = '\002J\001K',
	MicrosoftWordE186DialogInsertFormField = '\002J\001\343',
	MicrosoftWordE186DialogFormFieldOptions = '\002J\001a',
	MicrosoftWordE186DialogInsertCaption = '\002J\001e',
	MicrosoftWordE186DialogInsertAutoCaption = '\002J\001g',
	MicrosoftWordE186DialogInsertAddCaption = '\002J\001\222',
	MicrosoftWordE186DialogInsertCaptionNumbering = '\002J\001f',
	MicrosoftWordE186DialogInsertCrossReference = '\002J\001o',
	MicrosoftWordE186DialogToolsManageFields = '\002J\002w',
	MicrosoftWordE186DialogToolsAutoManager = '\002J\003\223',
	MicrosoftWordE186DialogToolsAutocorrect = '\002J\001z',
	MicrosoftWordE186DialogToolsAutocorrectExceptions = '\002J\002\372',
	MicrosoftWordE186DialogConnect = '\002J\001\244',
	MicrosoftWordE186DialogToolsOptionsView = '\002J\000\314',
	MicrosoftWordE186DialogInsertSubdocument = '\002J\002G',
	MicrosoftWordE186DialogFileRoutingSlip = '\002J\002p',
	MicrosoftWordE186DialogFontSubstitution = '\002J\002E',
	MicrosoftWordE186DialogToolsOptionsTypography = '\002J\002\343',
	MicrosoftWordE186DialogToolsOptionsAutoFormatAsYouType = '\002J\003\012',
	MicrosoftWordE186DialogControlRun = '\002J\000\353',
	MicrosoftWordE186DialogFileVersions = '\002J\003\261',
	MicrosoftWordE186DialogFileSaveVersion = '\002J\003\357',
	MicrosoftWordE186DialogWindowActivate = '\002J\000\334',
	MicrosoftWordE186DialogToolsMacroRecord = '\002J\000\326',
	MicrosoftWordE186DialogToolsRevisions = '\002J\000\305',
	MicrosoftWordE186DialogWebOptions = '\002J\003\202',
	MicrosoftWordE186DialogFitText = '\002J\003\327',
	MicrosoftWordE186DialogFormatEncloseCharacters = '\002J\004\212',
	MicrosoftWordE186DialogEmail = '\002J\020%',
	MicrosoftWordE186DialogFormatTheme = '\002J\003W',
	MicrosoftWordE186DialogToolsOptionsSecurity = '\002J\005Q',
	MicrosoftWordE186DialogToolsOptionsFeedback = '\002J\006\301',
	MicrosoftWordE186DialogToolsOptionsEditCopyPaste = '\002J\005L',
	MicrosoftWordE186DialogToolsOptionsNoteRecording = '\002J\006a',
	MicrosoftWordE186DialogMathRecognizedFunctions = '\002J\006\321',
	MicrosoftWordE186DialogMathMatrixSpacing = '\002J\006\322',
	MicrosoftWordE186DialogMathEquationArraySpacing = '\002J\006\323'
};
typedef enum MicrosoftWordE186 MicrosoftWordE186;

enum MicrosoftWordE187 {
	MicrosoftWordE187FieldKindNone = '\002K\000\000',
	MicrosoftWordE187FieldKindHot = '\002K\000\001',
	MicrosoftWordE187FieldKindWarm = '\002K\000\002',
	MicrosoftWordE187FieldKindCold = '\002K\000\003'
};
typedef enum MicrosoftWordE187 MicrosoftWordE187;

enum MicrosoftWordE188 {
	MicrosoftWordE188RegularText = '\002L\000\000',
	MicrosoftWordE188NumberText = '\002L\000\001',
	MicrosoftWordE188DateText = '\002L\000\002',
	MicrosoftWordE188CurrentDateText = '\002L\000\003',
	MicrosoftWordE188CurrentTimeText = '\002L\000\004',
	MicrosoftWordE188CalculationText = '\002L\000\005'
};
typedef enum MicrosoftWordE188 MicrosoftWordE188;

enum MicrosoftWordE189 {
	MicrosoftWordE189NeverConvert = '\002M\000\000',
	MicrosoftWordE189AlwaysConvert = '\002M\000\001',
	MicrosoftWordE189AskToNotConvert = '\002M\000\002',
	MicrosoftWordE189AskToConvert = '\002M\000\003'
};
typedef enum MicrosoftWordE189 MicrosoftWordE189;

enum MicrosoftWordE190 {
	MicrosoftWordE190NotAMergeDocument = '\002M\377\377',
	MicrosoftWordE190DocumentTypeFormLetters = '\002N\000\000',
	MicrosoftWordE190DocumentTypeMailingLabels = '\002N\000\001',
	MicrosoftWordE190DocumentTypeEnvelopes = '\002N\000\002',
	MicrosoftWordE190DocumentTypeCatalog = '\002N\000\003'
};
typedef enum MicrosoftWordE190 MicrosoftWordE190;

enum MicrosoftWordE191 {
	MicrosoftWordE191NormalDocument = '\002O\000\000',
	MicrosoftWordE191MainDocumentOnly = '\002O\000\001',
	MicrosoftWordE191MainAndDataSource = '\002O\000\002',
	MicrosoftWordE191MainAndHeader = '\002O\000\003',
	MicrosoftWordE191MainAndSourceAndHeader = '\002O\000\004',
	MicrosoftWordE191DataSource = '\002O\000\005'
};
typedef enum MicrosoftWordE191 MicrosoftWordE191;

enum MicrosoftWordE192 {
	MicrosoftWordE192SendToNewDocument = '\002P\000\000',
	MicrosoftWordE192SendToPrinter = '\002P\000\001',
	MicrosoftWordE192SendToEmail = '\002P\000\002',
	MicrosoftWordE192SendToFax = '\002P\000\003'
};
typedef enum MicrosoftWordE192 MicrosoftWordE192;

enum MicrosoftWordE193 {
	MicrosoftWordE193NoActiveRecord = '\002P\377\377',
	MicrosoftWordE193NextDataRecord = '\002P\377\376',
	MicrosoftWordE193PreviousDataRecord = '\002P\377\375',
	MicrosoftWordE193FirstDataRecord = '\002P\377\374',
	MicrosoftWordE193LastDataRecord = '\002P\377\373'
};
typedef enum MicrosoftWordE193 MicrosoftWordE193;

enum MicrosoftWordE194 {
	MicrosoftWordE194DefaultFirstRecord = '\000\000\000\001',
	MicrosoftWordE194DefaultLastRecord = '\377\377\377\360'
};
typedef enum MicrosoftWordE194 MicrosoftWordE194;

enum MicrosoftWordE195 {
	MicrosoftWordE195NoMergeInfo = '\002R\377\377',
	MicrosoftWordE195MergeInfoFromWord = '\002S\000\000',
	MicrosoftWordE195MergeInfoFromAccessDde = '\002S\000\001',
	MicrosoftWordE195MergeInfoFromExcelDde = '\002S\000\002',
	MicrosoftWordE195MergeInfoFromMsqueryDde = '\002S\000\003',
	MicrosoftWordE195MergeInfoFromOdbc = '\002S\000\004'
};
typedef enum MicrosoftWordE195 MicrosoftWordE195;

enum MicrosoftWordE196 {
	MicrosoftWordE196MergeIfEqual = '\002T\000\000',
	MicrosoftWordE196MergeIfNotEqual = '\002T\000\001',
	MicrosoftWordE196MergeIfLessThan = '\002T\000\002',
	MicrosoftWordE196MergeIfGreaterThan = '\002T\000\003',
	MicrosoftWordE196MergeIfLessThanOrEqual = '\002T\000\004',
	MicrosoftWordE196MergeIfGreaterThanOrEqual = '\002T\000\005',
	MicrosoftWordE196MergeIfIsBlank = '\002T\000\006',
	MicrosoftWordE196MergeIfIsNotBlank = '\002T\000\007'
};
typedef enum MicrosoftWordE196 MicrosoftWordE196;

enum MicrosoftWordE197 {
	MicrosoftWordE197SortByName = '\002U\000\000',
	MicrosoftWordE197SortByLocation = '\002U\000\001'
};
typedef enum MicrosoftWordE197 MicrosoftWordE197;

enum MicrosoftWordE198 {
	MicrosoftWordE198WindowStateNormal = '\002V\000\000',
	MicrosoftWordE198WindowStateMaximize = '\002V\000\001',
	MicrosoftWordE198WindowStateMinimize = '\002V\000\002'
};
typedef enum MicrosoftWordE198 MicrosoftWordE198;

enum MicrosoftWordE199 {
	MicrosoftWordE199LinkNone = '\002W\000\000',
	MicrosoftWordE199LinkDataInDoc = '\002W\000\001',
	MicrosoftWordE199LinkDataOnDisk = '\002W\000\002'
};
typedef enum MicrosoftWordE199 MicrosoftWordE199;

enum MicrosoftWordE200 {
	MicrosoftWordE200LinkTypeOle = '\002X\000\000',
	MicrosoftWordE200LinkTypePicture = '\002X\000\001',
	MicrosoftWordE200LinkTypeText = '\002X\000\002',
	MicrosoftWordE200LinkTypeReference = '\002X\000\003',
	MicrosoftWordE200LinkTypeInclude = '\002X\000\004',
	MicrosoftWordE200LinkTypeImport = '\002X\000\005',
	MicrosoftWordE200LinkTypeDde = '\002X\000\006',
	MicrosoftWordE200LinkTypeDdeauto = '\002X\000\007',
	MicrosoftWordE200LinkTypeChart = '\002X\000\010'
};
typedef enum MicrosoftWordE200 MicrosoftWordE200;

enum MicrosoftWordE201 {
	MicrosoftWordE201WindowDocument = '\002Y\000\000',
	MicrosoftWordE201WindowTemplate = '\002Y\000\001'
};
typedef enum MicrosoftWordE201 MicrosoftWordE201;

enum MicrosoftWordE202 {
	MicrosoftWordE202NormalView = '\002Z\000\001',
	MicrosoftWordE202DraftView = '\002Z\000\001',
	MicrosoftWordE202OutlineView = '\002Z\000\002',
	MicrosoftWordE202PageView = '\002Z\000\003',
	MicrosoftWordE202PrintView = '\002Z\000\003',
	MicrosoftWordE202PrintPreviewView = '\002Z\000\004',
	MicrosoftWordE202MasterView = '\002Z\000\005',
	MicrosoftWordE202OnlineView = '\002Z\000\006',
	MicrosoftWordE202WordNoteView = '\002Z\000\007',
	MicrosoftWordE202PublishingView = '\002Z\000\010',
	MicrosoftWordE202ConflictView = '\002Z\000\011'
};
typedef enum MicrosoftWordE202 MicrosoftWordE202;

enum MicrosoftWordE203 {
	MicrosoftWordE203SeekMainDocument = '\002[\000\000',
	MicrosoftWordE203SeekPrimaryHeader = '\002[\000\001',
	MicrosoftWordE203SeekFirstPageHeader = '\002[\000\002',
	MicrosoftWordE203SeekEvenPagesHeader = '\002[\000\003',
	MicrosoftWordE203SeekPrimaryFooter = '\002[\000\004',
	MicrosoftWordE203SeekFirstPageFooter = '\002[\000\005',
	MicrosoftWordE203SeekEvenPagesFooter = '\002[\000\006',
	MicrosoftWordE203SeekFootnotes = '\002[\000\007',
	MicrosoftWordE203SeekEndnotes = '\002[\000\010',
	MicrosoftWordE203SeekCurrentPageHeader = '\002[\000\011',
	MicrosoftWordE203SeekCurrentPageFooter = '\002[\000\012'
};
typedef enum MicrosoftWordE203 MicrosoftWordE203;

enum MicrosoftWordE204 {
	MicrosoftWordE204PaneNone = '\002\\\000\000',
	MicrosoftWordE204PanePrimaryHeader = '\002\\\000\001',
	MicrosoftWordE204PaneFirstPageHeader = '\002\\\000\002',
	MicrosoftWordE204PaneEvenPagesHeader = '\002\\\000\003',
	MicrosoftWordE204PanePrimaryFooter = '\002\\\000\004',
	MicrosoftWordE204PaneFirstPageFooter = '\002\\\000\005',
	MicrosoftWordE204PaneEvenPagesFooter = '\002\\\000\006',
	MicrosoftWordE204PaneFootnotes = '\002\\\000\007',
	MicrosoftWordE204PaneEndnotes = '\002\\\000\010',
	MicrosoftWordE204PaneFootnoteContinuationNotice = '\002\\\000\011',
	MicrosoftWordE204PaneFootnoteContinuationSeparator = '\002\\\000\012',
	MicrosoftWordE204PaneFootnoteSeparator = '\002\\\000\013',
	MicrosoftWordE204PaneEndnoteContinuationNotice = '\002\\\000\014',
	MicrosoftWordE204PaneEndnoteContinuationSeparator = '\002\\\000\015',
	MicrosoftWordE204PaneEndnoteSeparator = '\002\\\000\016',
	MicrosoftWordE204PaneComments = '\002\\\000\017',
	MicrosoftWordE204PaneCurrentPageHeader = '\002\\\000\020',
	MicrosoftWordE204PaneCurrentPageFooter = '\002\\\000\021',
	MicrosoftWordE204PaneRevisions = '\002\\\000\022'
};
typedef enum MicrosoftWordE204 MicrosoftWordE204;

enum MicrosoftWordE205 {
	MicrosoftWordE205PageFitNone = '\002]\000\000',
	MicrosoftWordE205PageFitFullPage = '\002]\000\001',
	MicrosoftWordE205PageFitBestFit = '\002]\000\002'
};
typedef enum MicrosoftWordE205 MicrosoftWordE205;

enum MicrosoftWordE206 {
	MicrosoftWordE206BrowsePage = '\002^\000\001',
	MicrosoftWordE206BrowseSection = '\002^\000\002',
	MicrosoftWordE206BrowseComment = '\002^\000\003',
	MicrosoftWordE206BrowseFootnote = '\002^\000\004',
	MicrosoftWordE206BrowseEndnote = '\002^\000\005',
	MicrosoftWordE206BrowseField = '\002^\000\006',
	MicrosoftWordE206BrowseTable = '\002^\000\007',
	MicrosoftWordE206BrowseGraphic = '\002^\000\010',
	MicrosoftWordE206BrowseHeading = '\002^\000\011',
	MicrosoftWordE206BrowseEdit = '\002^\000\012',
	MicrosoftWordE206BrowseFind = '\002^\000\013',
	MicrosoftWordE206BrowseGoTo = '\002^\000\014'
};
typedef enum MicrosoftWordE206 MicrosoftWordE206;

enum MicrosoftWordE207 {
	MicrosoftWordE207PrinterDefaultBin = '\002_\000\000',
	MicrosoftWordE207PrinterUpperBin = '\002_\000\001',
	MicrosoftWordE207PrinterOnlyBin = '\002_\000\001',
	MicrosoftWordE207PrinterLowerBin = '\002_\000\002',
	MicrosoftWordE207PrinterMiddleBin = '\002_\000\003',
	MicrosoftWordE207PrinterManualFeed = '\002_\000\004',
	MicrosoftWordE207PrinterEnvelopeFeed = '\002_\000\005',
	MicrosoftWordE207PrinterManualEnvelopeFeed = '\002_\000\006',
	MicrosoftWordE207PrinterAutomaticSheetFeed = '\002_\000\007',
	MicrosoftWordE207PrinterTractorFeed = '\002_\000\010',
	MicrosoftWordE207PrinterSmallFormatBin = '\002_\000\011',
	MicrosoftWordE207PrinterLargeFormatBin = '\002_\000\012',
	MicrosoftWordE207PrinterLargeCapacityBin = '\002_\000\013',
	MicrosoftWordE207PrinterPaperCassette = '\002_\000\016',
	MicrosoftWordE207PrinterFormSource = '\002_\000\017'
};
typedef enum MicrosoftWordE207 MicrosoftWordE207;

enum MicrosoftWordE208 {
	MicrosoftWordE208OrientPortrait = '\002`\000\000',
	MicrosoftWordE208OrientLandscape = '\002`\000\001'
};
typedef enum MicrosoftWordE208 MicrosoftWordE208;

enum MicrosoftWordE209 {
	MicrosoftWordE209NoSelection = '\002a\000\000',
	MicrosoftWordE209SelectionIp = '\002a\000\001',
	MicrosoftWordE209SelectionNormal = '\002a\000\002',
	MicrosoftWordE209SelectionFrame = '\002a\000\003',
	MicrosoftWordE209SelectionColumn = '\002a\000\004',
	MicrosoftWordE209SelectionRow = '\002a\000\005',
	MicrosoftWordE209SelectionBlock = '\002a\000\006',
	MicrosoftWordE209SelectionInlineShape = '\002a\000\007',
	MicrosoftWordE209SelectionShape = '\002a\000\010'
};
typedef enum MicrosoftWordE209 MicrosoftWordE209;

enum MicrosoftWordE210 {
	MicrosoftWordE210CaptionFigure = '\002a\377\377',
	MicrosoftWordE210CaptionTable = '\002a\377\376',
	MicrosoftWordE210CaptionEquation = '\002a\377\375'
};
typedef enum MicrosoftWordE210 MicrosoftWordE210;

enum MicrosoftWordE211 {
	MicrosoftWordE211ReferenceTypeNumberedItem = '\002c\000\000',
	MicrosoftWordE211ReferenceTypeHeading = '\002c\000\001',
	MicrosoftWordE211ReferenceTypeBookmark = '\002c\000\002',
	MicrosoftWordE211ReferenceTypeFootnote = '\002c\000\003',
	MicrosoftWordE211ReferenceTypeEndnote = '\002c\000\004'
};
typedef enum MicrosoftWordE211 MicrosoftWordE211;

enum MicrosoftWordE212 {
	MicrosoftWordE212ReferenceContentText = '\002c\377\377',
	MicrosoftWordE212ReferenceNumberRelativeContext = '\002c\377\376',
	MicrosoftWordE212ReferenceNumberNoContext = '\002c\377\375',
	MicrosoftWordE212ReferenceNumberFullContext = '\002c\377\374',
	MicrosoftWordE212ReferenceEntireCaption = '\002d\000\002',
	MicrosoftWordE212ReferenceOnlyLabelAndNumber = '\002d\000\003',
	MicrosoftWordE212ReferenceOnlyCaptionText = '\002d\000\004',
	MicrosoftWordE212ReferenceFootnoteNumber = '\002d\000\005',
	MicrosoftWordE212ReferenceEndnoteNumber = '\002d\000\006',
	MicrosoftWordE212ReferencePageNumber = '\002d\000\007',
	MicrosoftWordE212ReferencePosition = '\002d\000\017',
	MicrosoftWordE212ReferenceFootnoteNumberFormatted = '\002d\000\020',
	MicrosoftWordE212ReferenceEndnoteNumberFormatted = '\002d\000\021'
};
typedef enum MicrosoftWordE212 MicrosoftWordE212;

enum MicrosoftWordE213 {
	MicrosoftWordE213IndexTemplate = '\002e\000\000',
	MicrosoftWordE213IndexClassic = '\002e\000\001',
	MicrosoftWordE213IndexFancy = '\002e\000\002',
	MicrosoftWordE213IndexModern = '\002e\000\003',
	MicrosoftWordE213IndexBulleted = '\002e\000\004',
	MicrosoftWordE213IndexFormal = '\002e\000\005',
	MicrosoftWordE213IndexSimple = '\002e\000\006'
};
typedef enum MicrosoftWordE213 MicrosoftWordE213;

enum MicrosoftWordE214 {
	MicrosoftWordE214IndexIndent = '\002f\000\000',
	MicrosoftWordE214IndexRunin = '\002f\000\001'
};
typedef enum MicrosoftWordE214 MicrosoftWordE214;

enum MicrosoftWordE215 {
	MicrosoftWordE215WrapNever = '\002g\000\000',
	MicrosoftWordE215WrapAlways = '\002g\000\001',
	MicrosoftWordE215WrapAsk = '\002g\000\002'
};
typedef enum MicrosoftWordE215 MicrosoftWordE215;

enum MicrosoftWordE216 {
	MicrosoftWordE216NoRevision = '\002h\000\000',
	MicrosoftWordE216RevisionInsert = '\002h\000\001',
	MicrosoftWordE216RevisionDelete = '\002h\000\002',
	MicrosoftWordE216RevisionProperty = '\002h\000\003',
	MicrosoftWordE216RevisionParagraphNumber = '\002h\000\004',
	MicrosoftWordE216RevisionDisplayField = '\002h\000\005',
	MicrosoftWordE216RevisionReconcile = '\002h\000\006',
	MicrosoftWordE216RevisionConflict = '\002h\000\007',
	MicrosoftWordE216RevisionStyle = '\002h\000\010',
	MicrosoftWordE216RevisionReplace = '\002h\000\011',
	MicrosoftWordE216RevisionParagraphProperty = '\002h\000\012',
	MicrosoftWordE216RevisionTableProperty = '\002h\000\013',
	MicrosoftWordE216RevisionSectionProperty = '\002h\000\014',
	MicrosoftWordE216RevisionStyleDefinition = '\002h\000\015',
	MicrosoftWordE216RevisionMoveFrom = '\002h\000\016',
	MicrosoftWordE216RevisionMoveTo = '\002h\000\017',
	MicrosoftWordE216RevisionCellInsertion = '\002h\000\020',
	MicrosoftWordE216RevisionCellDeletion = '\002h\000\021',
	MicrosoftWordE216RevisionCellMerge = '\002h\000\022',
	MicrosoftWordE216RevisionCellSplit = '\002h\000\023',
	MicrosoftWordE216RevisionConflictInsert = '\002h\000\024',
	MicrosoftWordE216RevisionConflictDelete = '\002h\000\025'
};
typedef enum MicrosoftWordE216 MicrosoftWordE216;

enum MicrosoftWordE217 {
	MicrosoftWordE217OneAfterAnother = '\002i\000\000',
	MicrosoftWordE217AllAtOnce = '\002i\000\001'
};
typedef enum MicrosoftWordE217 MicrosoftWordE217;

enum MicrosoftWordE218 {
	MicrosoftWordE218NotYetRouted = '\002j\000\000',
	MicrosoftWordE218RouteInProgress = '\002j\000\001',
	MicrosoftWordE218RouteComplete = '\002j\000\002'
};
typedef enum MicrosoftWordE218 MicrosoftWordE218;

enum MicrosoftWordE219 {
	MicrosoftWordE219SectionContinuous = '\002k\000\000',
	MicrosoftWordE219SectionNewColumn = '\002k\000\001',
	MicrosoftWordE219SectionNewPage = '\002k\000\002',
	MicrosoftWordE219SectionEvenPage = '\002k\000\003',
	MicrosoftWordE219SectionOddPage = '\002k\000\004'
};
typedef enum MicrosoftWordE219 MicrosoftWordE219;

enum MicrosoftWordE220 {
	MicrosoftWordE220DoNotSaveChanges = '\002l\000\000',
	MicrosoftWordE220SaveChanges = '\002k\377\377',
	MicrosoftWordE220PromptToSaveChanges = '\002k\377\376'
};
typedef enum MicrosoftWordE220 MicrosoftWordE220;

enum MicrosoftWordE221 {
	MicrosoftWordE221DocumentNotSpecified = '\002m\000\000',
	MicrosoftWordE221DocumentLetter = '\002m\000\001',
	MicrosoftWordE221DocumentEmail = '\002m\000\002'
};
typedef enum MicrosoftWordE221 MicrosoftWordE221;

enum MicrosoftWordE222 {
	MicrosoftWordE222TypeDocument = '\002n\000\000',
	MicrosoftWordE222TypeTemplate = '\002n\000\001'
};
typedef enum MicrosoftWordE222 MicrosoftWordE222;

enum MicrosoftWordE223 {
	MicrosoftWordE223WordDocument = '\002o\000\000',
	MicrosoftWordE223OriginalDocumentFormat = '\002o\000\001',
	MicrosoftWordE223PromptUser = '\002o\000\002'
};
typedef enum MicrosoftWordE223 MicrosoftWordE223;

enum MicrosoftWordE224 {
	MicrosoftWordE224RelocateUp = '\002p\000\000',
	MicrosoftWordE224RelocateDown = '\002p\000\001'
};
typedef enum MicrosoftWordE224 MicrosoftWordE224;

enum MicrosoftWordE225 {
	MicrosoftWordE225InsertedTextMarkNone = '\002q\000\000',
	MicrosoftWordE225InsertedTextMarkBold = '\002q\000\001',
	MicrosoftWordE225InsertedTextMarkItalic = '\002q\000\002',
	MicrosoftWordE225InsertedTextMarkUnderline = '\002q\000\003',
	MicrosoftWordE225InsertedTextMarkDoubleUnderline = '\002q\000\004',
	MicrosoftWordE225InsertedTextMarkColorOnly = '\002q\000\005',
	MicrosoftWordE225InsertedTextMarkStrikeThrough = '\002q\000\006',
	MicrosoftWordE225InsertedTextMarkDoubleStrikeThrough = '\002q\000\007'
};
typedef enum MicrosoftWordE225 MicrosoftWordE225;

enum MicrosoftWordE226 {
	MicrosoftWordE226RevisedLinesMarkNone = '\002r\000\000',
	MicrosoftWordE226RevisedLinesMarkLeftBorder = '\002r\000\001',
	MicrosoftWordE226RevisedLinesMarkRightBorder = '\002r\000\002',
	MicrosoftWordE226RevisedLinesMarkOutsideBorder = '\002r\000\003'
};
typedef enum MicrosoftWordE226 MicrosoftWordE226;

enum MicrosoftWordE227 {
	MicrosoftWordE227DeletedTextMarkHidden = '\002s\000\000',
	MicrosoftWordE227DeletedTextMarkStrikeThrough = '\002s\000\001',
	MicrosoftWordE227DeletedTextMarkCaret = '\002s\000\002',
	MicrosoftWordE227DeletedTextMarkPound = '\002s\000\003',
	MicrosoftWordE227DeletedTextMarkNone = '\002s\000\004',
	MicrosoftWordE227DeletedTextMarkBold = '\002s\000\005',
	MicrosoftWordE227DeletedTextMarkItalic = '\002s\000\006',
	MicrosoftWordE227DeletedTextMarkUnderline = '\002s\000\007',
	MicrosoftWordE227DeletedTextMarkDoubleUnderline = '\002s\000\010',
	MicrosoftWordE227DeletedTextMarkColorOnly = '\002s\000\011',
	MicrosoftWordE227DeletedTextMarkDoubleStrikeThrough = '\002s\000\012'
};
typedef enum MicrosoftWordE227 MicrosoftWordE227;

enum MicrosoftWordE228 {
	MicrosoftWordE228RevisedPropertiesMarkNone = '\002t\000\000',
	MicrosoftWordE228RevisedPropertiesMarkBold = '\002t\000\001',
	MicrosoftWordE228RevisedPropertiesMarkItalic = '\002t\000\002',
	MicrosoftWordE228RevisedPropertiesMarkUnderline = '\002t\000\003',
	MicrosoftWordE228RevisedPropertiesMarkDoubleUnderline = '\002t\000\004',
	MicrosoftWordE228RevisedPropertiesMarkColorOnly = '\002t\000\005',
	MicrosoftWordE228RevisedPropertiesMarkStrikeThrough = '\002t\000\006',
	MicrosoftWordE228RevisedPropertiesMarkDoubleStrikeThrough = '\002t\000\007'
};
typedef enum MicrosoftWordE228 MicrosoftWordE228;

enum MicrosoftWordE316 {
	MicrosoftWordE316MoveToTextMarkNone = '\002\314\000\000',
	MicrosoftWordE316MoveToTextMarkBold = '\002\314\000\001',
	MicrosoftWordE316MoveToTextMarkItalic = '\002\314\000\002',
	MicrosoftWordE316MoveToTextMarkUnderline = '\002\314\000\003',
	MicrosoftWordE316MoveToTextMarkDoubleUnderline = '\002\314\000\004',
	MicrosoftWordE316MoveToTextMarkColorOnly = '\002\314\000\005',
	MicrosoftWordE316MoveToTextMarkStrikeThrough = '\002\314\000\006',
	MicrosoftWordE316MoveToTextMarkDoubleStrikeThrough = '\002\314\000\007'
};
typedef enum MicrosoftWordE316 MicrosoftWordE316;

enum MicrosoftWordE319 {
	MicrosoftWordE319MathAccentType = '\002\317\000\001',
	MicrosoftWordE319MathFunctionBarType = '\002\317\000\002',
	MicrosoftWordE319MathBoxType = '\002\317\000\003',
	MicrosoftWordE319MathBoxedFormulaType = '\002\317\000\004',
	MicrosoftWordE319MathDelimitersType = '\002\317\000\005',
	MicrosoftWordE319MathEquationArrayType = '\002\317\000\006',
	MicrosoftWordE319MathFractionType = '\002\317\000\007',
	MicrosoftWordE319MathFunctionApplyType = '\002\317\000\010',
	MicrosoftWordE319MathStretchStackType = '\002\317\000\011',
	MicrosoftWordE319MathLowerLimitType = '\002\317\000\012',
	MicrosoftWordE319MathUpperLimitType = '\002\317\000\013',
	MicrosoftWordE319MathMatrixType = '\002\317\000\014',
	MicrosoftWordE319MathNaryOperatorType = '\002\317\000\015',
	MicrosoftWordE319MathPhantomType = '\002\317\000\016',
	MicrosoftWordE319MathLeftSubSuperscriptType = '\002\317\000\017',
	MicrosoftWordE319MathRadicalType = '\002\317\000\020',
	MicrosoftWordE319MathSubscriptType = '\002\317\000\021',
	MicrosoftWordE319MathSubSuperscriptType = '\002\317\000\022',
	MicrosoftWordE319MathSuperscriptType = '\002\317\000\023',
	MicrosoftWordE319MathTextType = '\002\317\000\024',
	MicrosoftWordE319MathNormalTextType = '\002\317\000\025',
	MicrosoftWordE319MathLiteralTextType = '\002\317\000\026'
};
typedef enum MicrosoftWordE319 MicrosoftWordE319;

enum MicrosoftWordE320 {
	MicrosoftWordE320EquationHorizontalAlignCenter = '\002\320\000\000',
	MicrosoftWordE320EquationHorizontalAlignLeft = '\002\320\000\001',
	MicrosoftWordE320EquationHorizontalAlignRight = '\002\320\000\002'
};
typedef enum MicrosoftWordE320 MicrosoftWordE320;

enum MicrosoftWordE321 {
	MicrosoftWordE321EquationVerticalAlignCenter = '\002\321\000\000',
	MicrosoftWordE321EquationVerticalAlignTop = '\002\321\000\001',
	MicrosoftWordE321EquationVerticalAlignBottom = '\002\321\000\002'
};
typedef enum MicrosoftWordE321 MicrosoftWordE321;

enum MicrosoftWordE322 {
	MicrosoftWordE322MathFractionTypeBar = '\002\322\000\000',
	MicrosoftWordE322MathFractionTypeStack = '\002\322\000\001',
	MicrosoftWordE322MathFractionTypeSlashed = '\002\322\000\002',
	MicrosoftWordE322MathFractionTypeLinear = '\002\322\000\003'
};
typedef enum MicrosoftWordE322 MicrosoftWordE322;

enum MicrosoftWordE323 {
	MicrosoftWordE323EquationSpacingSingle = '\002\323\000\000',
	MicrosoftWordE323EquationSpacingOneAndAHalf = '\002\323\000\001',
	MicrosoftWordE323EquationSpacingDouble = '\002\323\000\002',
	MicrosoftWordE323EquationSpacingExactly = '\002\323\000\003',
	MicrosoftWordE323EquationSpacingMultiple = '\002\323\000\004'
};
typedef enum MicrosoftWordE323 MicrosoftWordE323;

enum MicrosoftWordE324 {
	MicrosoftWordE324EquationDisplayProfessional = '\002\324\000\000' /* Specifies an equation is displayed in professional/built up format. */,
	MicrosoftWordE324EquationDisplayInline = '\002\324\000\001' /* Specifies an equation is displayed in linear/built down style. */
};
typedef enum MicrosoftWordE324 MicrosoftWordE324;

enum MicrosoftWordE325 {
	MicrosoftWordE325MathDelimitersCenterVertically = '\002\325\000\000',
	MicrosoftWordE325MathDelimitersMatchContentSize = '\002\325\000\001'
};
typedef enum MicrosoftWordE325 MicrosoftWordE325;

enum MicrosoftWordE326 {
	MicrosoftWordE326EquationJustificationCenterGroup = '\002\326\000\001',
	MicrosoftWordE326EquationJustificationCenter = '\002\326\000\002',
	MicrosoftWordE326EquationJustificationLeft = '\002\326\000\003',
	MicrosoftWordE326EquationJustificationRight = '\002\326\000\004',
	MicrosoftWordE326EquationJustificationInline = '\002\326\000\007'
};
typedef enum MicrosoftWordE326 MicrosoftWordE326;

enum MicrosoftWordE327 {
	MicrosoftWordE327MathBinaryOperatorBreakBefore = '\002\327\000\000',
	MicrosoftWordE327MathBinaryOperatorBreakAfter = '\002\327\000\001',
	MicrosoftWordE327MathBinaryOperatorRepeat = '\002\327\000\002'
};
typedef enum MicrosoftWordE327 MicrosoftWordE327;

enum MicrosoftWordE328 {
	MicrosoftWordE328MinusMinus = '\002\330\000\000',
	MicrosoftWordE328PlusMinus = '\002\330\000\001',
	MicrosoftWordE328MinusPlus = '\002\330\000\002'
};
typedef enum MicrosoftWordE328 MicrosoftWordE328;

enum MicrosoftWordE329 {
	MicrosoftWordE329BuildingBlockEquations = '\000\000\000\003'
};
typedef enum MicrosoftWordE329 MicrosoftWordE329;

enum MicrosoftWordE330 {
	MicrosoftWordE330InlineBuildingBlock = '\000\000\000\000',
	MicrosoftWordE330PageLevelBuildingBlock = '\000\000\000\001',
	MicrosoftWordE330ParagraphLevelBuildingBlock = '\000\000\000\002'
};
typedef enum MicrosoftWordE330 MicrosoftWordE330;

enum MicrosoftWordE317 {
	MicrosoftWordE317MoveFromTextMarkHidden = '\002\315\000\000',
	MicrosoftWordE317MoveFromTextMarkDoubleStrikeThrough = '\002\315\000\001',
	MicrosoftWordE317MoveFromTextMarkStrikeThrough = '\002\315\000\002',
	MicrosoftWordE317MoveFromTextMarkCaret = '\002\315\000\003',
	MicrosoftWordE317MoveFromTextMarkPound = '\002\315\000\004',
	MicrosoftWordE317MoveFromTextMarkNone = '\002\315\000\005',
	MicrosoftWordE317MoveFromTextMarkBold = '\002\315\000\006',
	MicrosoftWordE317MoveFromTextMarkItalic = '\002\315\000\007',
	MicrosoftWordE317MoveFromTextMarkUnderline = '\002\315\000\010',
	MicrosoftWordE317MoveFromTextMarkDoubleUnderline = '\002\315\000\011',
	MicrosoftWordE317MoveFromTextMarkColorOnly = '\002\315\000\012'
};
typedef enum MicrosoftWordE317 MicrosoftWordE317;

enum MicrosoftWordE318 {
	MicrosoftWordE318CellColorByAuthor = '\002\315\377\377',
	MicrosoftWordE318CellColorNoHighlight = '\002\316\000\000',
	MicrosoftWordE318CellColorPink = '\002\316\000\001',
	MicrosoftWordE318CellColorLightBlue = '\002\316\000\002',
	MicrosoftWordE318CellColorLightYellow = '\002\316\000\003',
	MicrosoftWordE318CellColorLightPurple = '\002\316\000\004',
	MicrosoftWordE318CellColorLightOrange = '\002\316\000\005',
	MicrosoftWordE318CellColorLightGreen = '\002\316\000\006',
	MicrosoftWordE318CellColorLightGray = '\002\316\000\007'
};
typedef enum MicrosoftWordE318 MicrosoftWordE318;

enum MicrosoftWordE229 {
	MicrosoftWordE229FieldShadingNever = '\002u\000\000',
	MicrosoftWordE229FieldShadingAlways = '\002u\000\001',
	MicrosoftWordE229FieldShadingWhenSelected = '\002u\000\002'
};
typedef enum MicrosoftWordE229 MicrosoftWordE229;

enum MicrosoftWordE230 {
	MicrosoftWordE230DocumentsPath = '\002v\000\000',
	MicrosoftWordE230PicturesPath = '\002v\000\001',
	MicrosoftWordE230UserTemplatesPath = '\002v\000\002',
	MicrosoftWordE230WorkgroupTemplatesPath = '\002v\000\003',
	MicrosoftWordE230UserOptionsPath = '\002v\000\004',
	MicrosoftWordE230AutoRecoverPath = '\002v\000\005',
	MicrosoftWordE230ToolsPath = '\002v\000\006',
	MicrosoftWordE230TutorialPath = '\002v\000\007',
	MicrosoftWordE230StartupPath = '\002v\000\010',
	MicrosoftWordE230ProgramPath = '\002v\000\011',
	MicrosoftWordE230GraphicsFiltersPath = '\002v\000\012',
	MicrosoftWordE230TextConvertersPath = '\002v\000\013',
	MicrosoftWordE230ProofingToolsPath = '\002v\000\014',
	MicrosoftWordE230TempFilePath = '\002v\000\015',
	MicrosoftWordE230CurrentFolderPath = '\002v\000\016',
	MicrosoftWordE230StyleGalleryPath = '\002v\000\017',
	MicrosoftWordE230TrashPath = '\002v\000\020',
	MicrosoftWordE230OfficePath = '\002v\000\021',
	MicrosoftWordE230TypeLibrariesPath = '\002v\000\022',
	MicrosoftWordE230BorderArtPath = '\002v\000\023'
};
typedef enum MicrosoftWordE230 MicrosoftWordE230;

enum MicrosoftWordE231 {
	MicrosoftWordE231NoTabHangingIndent = '\002w\000\001',
	MicrosoftWordE231NoSpaceForRaisedOrLoweredCharacters = '\002w\000\002',
	MicrosoftWordE231PrintColorsBlack = '\002w\000\003',
	MicrosoftWordE231WrapTrailSpaces = '\002w\000\004',
	MicrosoftWordE231NoColumnBalance = '\002w\000\005',
	MicrosoftWordE231ConvertDataMergeEscapes = '\002w\000\006',
	MicrosoftWordE231SuppressSpaceBeforeAfterPageBreak = '\002w\000\007',
	MicrosoftWordE231SuppressTopSpacing = '\002w\000\010',
	MicrosoftWordE231OriginalWordTableRules = '\002w\000\011',
	MicrosoftWordE231TransparentMetafiles = '\002w\000\012',
	MicrosoftWordE231ShowBreaksInFrames = '\002w\000\013',
	MicrosoftWordE231SwapBordersFacingPages = '\002w\000\014',
	MicrosoftWordE231LeaveBackslashAlone = '\002w\000\015',
	MicrosoftWordE231ExpandShiftReturn = '\002w\000\016',
	MicrosoftWordE231DoNotUnderlineTrailingSpaces = '\002w\000\017',
	MicrosoftWordE231DoNotBalanceSBCSAndDBCSCharacters = '\002w\000\020',
	MicrosoftWordE231SuppressTopSpacingMacWord5 = '\002w\000\021',
	MicrosoftWordE231SpacingInWholePoints = '\002w\000\022',
	MicrosoftWordE231PrintBodyTextBeforeHeader = '\002w\000\023',
	MicrosoftWordE231NoExtraSpaceBetweenRowsOfText = '\002w\000\024',
	MicrosoftWordE231NoSpaceForUnderlines = '\002w\000\025',
	MicrosoftWordE231UseLargerSmallCaps = '\002w\000\026',
	MicrosoftWordE231NoExtraLineSpacing = '\002w\000\027',
	MicrosoftWordE231TruncateFontHeight = '\002w\000\030',
	MicrosoftWordE231SubstituteFontBySize = '\002w\000\031',
	MicrosoftWordE231UsePrinterMetrics = '\002w\000\032',
	MicrosoftWordE231Word6BorderRules = '\002w\000\033',
	MicrosoftWordE231ExactOnTop = '\002w\000\034',
	MicrosoftWordE231SuppressBottomSpacing = '\002w\000\035',
	MicrosoftWordE231WordPerfectSpaceWidth = '\002w\000\036',
	MicrosoftWordE231WordPerfectJustification = '\002w\000\037',
	MicrosoftWordE231Word6LineWrap = '\002w\000 ',
	MicrosoftWordE231Word96ShapeLayout = '\002w\000!',
	MicrosoftWordE231Word98FootnoteLayout = '\002w\000\"',
	MicrosoftWordE231DoNotUseHtmlParagraphAutoSpacing = '\002w\000#',
	MicrosoftWordE231DoNotAdjustLineHeightInTable = '\002w\000$',
	MicrosoftWordE231ForgetLastTabAlignment = '\002w\000%',
	MicrosoftWordE231Word95AutoSpace = '\002w\000&',
	MicrosoftWordE231AlignTablesRowByRow = '\002w\000\'',
	MicrosoftWordE231LayoutRawTableWidth = '\002w\000(',
	MicrosoftWordE231LayoutTableRowsApart = '\002w\000)',
	MicrosoftWordE231UseWord97LineBreakingRules = '\002w\000*',
	MicrosoftWordE231DoNotBreakWrappedTables = '\002w\000+',
	MicrosoftWordE231GrowAutofit = '\002w\000,',
	MicrosoftWordE231DoNotSnapTextToGridInTableWithObjects = '\002w\000-',
	MicrosoftWordE231SelectFieldWithFirstOrLastCharacter = '\002w\000.',
	MicrosoftWordE231ApplyBreakingRules = '\002w\000/',
	MicrosoftWordE231DoNotWrapTextWithPunctuation = '\002w\0000',
	MicrosoftWordE231DoNotUseAsianBreakRulesInGrid = '\002w\0001',
	MicrosoftWordE231UseWord2002TableStyleRules = '\002w\0002'
};
typedef enum MicrosoftWordE231 MicrosoftWordE231;

enum MicrosoftWordE314 {
	MicrosoftWordE314DefaultCompatibilitySettings = '\002\312\000\000' /* Microsoft Word 2007-2008 */,
	MicrosoftWordE314Word20002004AndX = '\002\312\000\014' /* Microsoft Word 2000-2004 and X */,
	MicrosoftWordE314Word97And98 = '\002\312\000\001' /* Microsoft Word 97-98 */,
	MicrosoftWordE314Word6And95 = '\002\312\000\002' /* Microsoft Word 6.0/95 */,
	MicrosoftWordE314WinWord1 = '\002\312\000\003' /* Word for Windows 1.0 */,
	MicrosoftWordE314WinWord2 = '\002\312\000\004' /* Word for Windows 2.0 */,
	MicrosoftWordE314MacWord5 = '\002\312\000\005' /* Word for the Macintosh 5.x */,
	MicrosoftWordE314DosWord = '\002\312\000\006' /* Word for MS-DOS */,
	MicrosoftWordE314Wordperfect5 = '\002\312\000\007' /* WordPerfect 5.x */,
	MicrosoftWordE314WinWordperfect6 = '\002\312\000\010' /* WordPerfect 6.x for Windows */,
	MicrosoftWordE314DosWordperfect6 = '\002\312\000\011' /* WordPerfect 6.0 for DOS */,
	MicrosoftWordE314AsianWord97And98 = '\002\312\000\012' /* Microsoft Word Asian Versions 97-98 */,
	MicrosoftWordE314UsWord6And95 = '\002\312\000\013' /* US Microsoft Word for Windows 6.0/95 */,
	MicrosoftWordE314CustomCompatibilitySettings = '\002\312\000\015' /* Custom settings */
};
typedef enum MicrosoftWordE314 MicrosoftWordE314;

enum MicrosoftWordE232 {
	MicrosoftWordE232PaperTenXFourteen = '\002x\000\000',
	MicrosoftWordE232PaperElevenXSeventeen = '\002x\000\001',
	MicrosoftWordE232PaperLetter = '\002x\000\002',
	MicrosoftWordE232PaperLetterSmall = '\002x\000\003',
	MicrosoftWordE232PaperLegal = '\002x\000\004',
	MicrosoftWordE232PaperExecutive = '\002x\000\005',
	MicrosoftWordE232PaperA3 = '\002x\000\006',
	MicrosoftWordE232PaperA4 = '\002x\000\007',
	MicrosoftWordE232PaperA4Small = '\002x\000\010',
	MicrosoftWordE232PaperA5 = '\002x\000\011',
	MicrosoftWordE232PaperB4 = '\002x\000\012',
	MicrosoftWordE232PaperB5 = '\002x\000\013',
	MicrosoftWordE232PaperCsheet = '\002x\000\014',
	MicrosoftWordE232PaperDsheet = '\002x\000\015',
	MicrosoftWordE232PaperEsheet = '\002x\000\016',
	MicrosoftWordE232PaperFanfoldLegalGerman = '\002x\000\017',
	MicrosoftWordE232PaperFanfoldStdGerman = '\002x\000\020',
	MicrosoftWordE232PaperFanfoldUs = '\002x\000\021',
	MicrosoftWordE232PaperFolio = '\002x\000\022',
	MicrosoftWordE232PaperLedger = '\002x\000\023',
	MicrosoftWordE232PaperNote = '\002x\000\024',
	MicrosoftWordE232PaperQuarto = '\002x\000\025',
	MicrosoftWordE232PaperStatement = '\002x\000\026',
	MicrosoftWordE232PaperTabloid = '\002x\000\027',
	MicrosoftWordE232PaperEnvelope9 = '\002x\000\030',
	MicrosoftWordE232PaperEnvelope10 = '\002x\000\031',
	MicrosoftWordE232PaperEnvelope11 = '\002x\000\032',
	MicrosoftWordE232PaperEnvelope12 = '\002x\000\033',
	MicrosoftWordE232PaperEnvelope14 = '\002x\000\034',
	MicrosoftWordE232PaperEnvelopeB4 = '\002x\000\035',
	MicrosoftWordE232PaperEnvelopeB5 = '\002x\000\036',
	MicrosoftWordE232PaperEnvelopeB6 = '\002x\000\037',
	MicrosoftWordE232PaperEnvelopeC3 = '\002x\000 ',
	MicrosoftWordE232PaperEnvelopeC4 = '\002x\000!',
	MicrosoftWordE232PaperEnvelopeC5 = '\002x\000\"',
	MicrosoftWordE232PaperEnvelopeC6 = '\002x\000#',
	MicrosoftWordE232PaperEnvelopeC65 = '\002x\000$',
	MicrosoftWordE232PaperEnvelopeDl = '\002x\000%',
	MicrosoftWordE232PaperEnvelopeItaly = '\002x\000&',
	MicrosoftWordE232PaperEnvelopeMonarch = '\002x\000\'',
	MicrosoftWordE232PaperEnvelopePersonal = '\002x\000(',
	MicrosoftWordE232PaperCustom = '\002x\000)'
};
typedef enum MicrosoftWordE232 MicrosoftWordE232;

enum MicrosoftWordE233 {
	MicrosoftWordE233CustomLabelLetter = '\002y\000\000',
	MicrosoftWordE233CustomLabelLetterLandscape = '\002y\000\001',
	MicrosoftWordE233CustomLabelA4 = '\002y\000\002',
	MicrosoftWordE233CustomLabelA4Landscape = '\002y\000\003',
	MicrosoftWordE233CustomLabelA5 = '\002y\000\004',
	MicrosoftWordE233CustomLabelA5Landscape = '\002y\000\005',
	MicrosoftWordE233CustomLabelB5 = '\002y\000\006',
	MicrosoftWordE233CustomLabelMini = '\002y\000\007',
	MicrosoftWordE233CustomLabelFanfold = '\002y\000\010'
};
typedef enum MicrosoftWordE233 MicrosoftWordE233;

enum MicrosoftWordE234 {
	MicrosoftWordE234NoDocumentProtection = '\002y\377\377',
	MicrosoftWordE234AllowOnlyRevisions = '\002z\000\000',
	MicrosoftWordE234AllowOnlyComments = '\002z\000\001',
	MicrosoftWordE234AllowOnlyFormFields = '\002z\000\002',
	MicrosoftWordE234AllowOnlyReading = '\002z\000\003'
};
typedef enum MicrosoftWordE234 MicrosoftWordE234;

enum MicrosoftWordE235 {
	MicrosoftWordE235Adjective = '\002{\000\000',
	MicrosoftWordE235Noun = '\002{\000\001',
	MicrosoftWordE235Adverb = '\002{\000\002',
	MicrosoftWordE235Verb = '\002{\000\003',
	MicrosoftWordE235Pronoun = '\002{\000\004',
	MicrosoftWordE235Conjunction = '\002{\000\005',
	MicrosoftWordE235Preposition = '\002{\000\006',
	MicrosoftWordE235Interjection = '\002{\000\007',
	MicrosoftWordE235Idiom = '\002{\000\010',
	MicrosoftWordE235Other = '\002{\000\011'
};
typedef enum MicrosoftWordE235 MicrosoftWordE235;

enum MicrosoftWordE236 {
	MicrosoftWordE236RelativeHorizontalPositionMargin = '\002|\000\000',
	MicrosoftWordE236RelativeHorizontalPositionPage = '\002|\000\001',
	MicrosoftWordE236RelativeHorizontalPositionColumn = '\002|\000\002',
	MicrosoftWordE236RelativeHorizontalPositionCharacter = '\002|\000\003',
	MicrosoftWordE236RelativeHorizontalPositionLeftMargin = '\002|\000\004',
	MicrosoftWordE236RelativeHorizontalPositionRightMargin = '\002|\000\005',
	MicrosoftWordE236RelativeHorizontalPositionInnerMargin = '\002|\000\006',
	MicrosoftWordE236RelativeHorizontalPositionOuterMargin = '\002|\000\007'
};
typedef enum MicrosoftWordE236 MicrosoftWordE236;

enum MicrosoftWordE237 {
	MicrosoftWordE237RelativeVerticalPositionMargin = '\002}\000\000',
	MicrosoftWordE237RelativeVerticalPositionPage = '\002}\000\001',
	MicrosoftWordE237RelativeVerticalPositionParagraph = '\002}\000\002',
	MicrosoftWordE237RelativeVerticalPositionLine = '\002}\000\003',
	MicrosoftWordE237RelativeVerticalPositionTopMargin = '\002}\000\004',
	MicrosoftWordE237RelativeVerticalPositionBottomMargin = '\002}\000\005',
	MicrosoftWordE237RelativeVerticalPositionInnerMargin = '\002}\000\006',
	MicrosoftWordE237RelativeVerticalPositionOuterMargin = '\002}\000\007'
};
typedef enum MicrosoftWordE237 MicrosoftWordE237;

enum MicrosoftWordE238 {
	MicrosoftWordE238Help = '\002~\000\000',
	MicrosoftWordE238HelpAbout = '\002~\000\001',
	MicrosoftWordE238HelpContents = '\002~\000\003',
	MicrosoftWordE238HelpIndex = '\002~\000\005',
	MicrosoftWordE238HelpPsshelp = '\002~\000\007',
	MicrosoftWordE238HelpSearch = '\002~\000\011'
};
typedef enum MicrosoftWordE238 MicrosoftWordE238;

enum MicrosoftWordE239 {
	MicrosoftWordE239KeyCategoryNil = '\002~\377\377',
	MicrosoftWordE239KeyCategoryDisable = '\002\177\000\000',
	MicrosoftWordE239KeyCategoryCommand = '\002\177\000\001',
	MicrosoftWordE239KeyCategoryMacro = '\002\177\000\002',
	MicrosoftWordE239KeyCategoryFont = '\002\177\000\003',
	MicrosoftWordE239KeyCategoryAutoText = '\002\177\000\004',
	MicrosoftWordE239KeyCategoryStyle = '\002\177\000\005',
	MicrosoftWordE239KeyCategorySymbol = '\002\177\000\006',
	MicrosoftWordE239KeyCategoryPrefix = '\002\177\000\007'
};
typedef enum MicrosoftWordE239 MicrosoftWordE239;

enum MicrosoftWordE240 {
	MicrosoftWordE240No_key = '\002\200\000\377',
	MicrosoftWordE240Shift_key = '\002\200\002\000',
	MicrosoftWordE240Control_key = '\002\200\020\000',
	MicrosoftWordE240Command_key = '\002\200\001\000',
	MicrosoftWordE240Option_key = '\002\200\010\000',
	MicrosoftWordE240KeyAlt = '\002\200\010\000',
	MicrosoftWordE240A_key = '\002\200\000A',
	MicrosoftWordE240B_key = '\002\200\000B',
	MicrosoftWordE240C_key = '\002\200\000C',
	MicrosoftWordE240D_key = '\002\200\000D',
	MicrosoftWordE240E_key = '\002\200\000E',
	MicrosoftWordE240F_key = '\002\200\000F',
	MicrosoftWordE240G_key = '\002\200\000G',
	MicrosoftWordE240H_key = '\002\200\000H',
	MicrosoftWordE240I_Key = '\002\200\000I',
	MicrosoftWordE240J_key = '\002\200\000J',
	MicrosoftWordE240K_key = '\002\200\000K',
	MicrosoftWordE240L_key = '\002\200\000L',
	MicrosoftWordE240M_key = '\002\200\000M',
	MicrosoftWordE240N_key = '\002\200\000N',
	MicrosoftWordE240O_key = '\002\200\000O',
	MicrosoftWordE240P_key = '\002\200\000P',
	MicrosoftWordE240Q_key = '\002\200\000Q',
	MicrosoftWordE240R_key = '\002\200\000R',
	MicrosoftWordE240S_key = '\002\200\000S',
	MicrosoftWordE240T_key = '\002\200\000T',
	MicrosoftWordE240U_kwy = '\002\200\000U',
	MicrosoftWordE240V_key = '\002\200\000V',
	MicrosoftWordE240W_key = '\002\200\000W',
	MicrosoftWordE240X_key = '\002\200\000X',
	MicrosoftWordE240Y_key = '\002\200\000Y',
	MicrosoftWordE240Z_key = '\002\200\000Z',
	MicrosoftWordE240Key_number_0 = '\002\200\0000',
	MicrosoftWordE240Key_number_1 = '\002\200\0001',
	MicrosoftWordE240Key_numbe_2 = '\002\200\0002',
	MicrosoftWordE240Key_numbe_3 = '\002\200\0003',
	MicrosoftWordE240Key_number_4 = '\002\200\0004',
	MicrosoftWordE240Key_number_5 = '\002\200\0005',
	MicrosoftWordE240Key_number_6 = '\002\200\0006',
	MicrosoftWordE240KeyNumber_7 = '\002\200\0007',
	MicrosoftWordE240Key_number_8 = '\002\200\0008',
	MicrosoftWordE240Key_number_9 = '\002\200\0009',
	MicrosoftWordE240Backspace_key = '\002\200\000\010',
	MicrosoftWordE240Tab_key = '\002\200\000\011',
	MicrosoftWordE240Key_numeric_5_special = '\002\200\000\014',
	MicrosoftWordE240Return_key = '\002\200\000\015',
	MicrosoftWordE240Pause_key = '\002\200\000\023',
	MicrosoftWordE240Esc_key = '\002\200\000\033',
	MicrosoftWordE240Spacebar_key = '\002\200\000 ',
	MicrosoftWordE240Page_up_key = '\002\200\000!',
	MicrosoftWordE240Page_down_key = '\002\200\000\"',
	MicrosoftWordE240End_key = '\002\200\000#',
	MicrosoftWordE240Home_key = '\002\200\000$',
	MicrosoftWordE240Insert_key = '\002\200\000-',
	MicrosoftWordE240Delete_key = '\002\200\000.',
	MicrosoftWordE240Key_numeric_0 = '\002\200\000`',
	MicrosoftWordE240Key_numeric_1 = '\002\200\000a',
	MicrosoftWordE240Key_numeric_2 = '\002\200\000b',
	MicrosoftWordE240Key_numeric_3 = '\002\200\000c',
	MicrosoftWordE240Key_numeric_4 = '\002\200\000d',
	MicrosoftWordE240Key_numeric_5 = '\002\200\000e',
	MicrosoftWordE240Key_numeric_6 = '\002\200\000f',
	MicrosoftWordE240Key_numeric_7 = '\002\200\000g',
	MicrosoftWordE240Key_numeric_8 = '\002\200\000h',
	MicrosoftWordE240Key_numeric_9 = '\002\200\000i',
	MicrosoftWordE240Key_numeric_multiply = '\002\200\000j',
	MicrosoftWordE240Key_numeric_add = '\002\200\000k',
	MicrosoftWordE240Key_numeric_subtract = '\002\200\000m',
	MicrosoftWordE240Key_numeric_decimal = '\002\200\000n',
	MicrosoftWordE240Key_numeric_divide = '\002\200\000o',
	MicrosoftWordE240F1_key = '\002\200\000p',
	MicrosoftWordE240F2_key = '\002\200\000q',
	MicrosoftWordE240F3_key = '\002\200\000r',
	MicrosoftWordE240F4_key = '\002\200\000s',
	MicrosoftWordE240F5_key = '\002\200\000t',
	MicrosoftWordE240F6_key = '\002\200\000u',
	MicrosoftWordE240F7_key = '\002\200\000v',
	MicrosoftWordE240F8_key = '\002\200\000w',
	MicrosoftWordE240F9_key = '\002\200\000x',
	MicrosoftWordE240F10_key = '\002\200\000y',
	MicrosoftWordE240F11_key = '\002\200\000z',
	MicrosoftWordE240F12_key = '\002\200\000{',
	MicrosoftWordE240F13_key = '\002\200\000|',
	MicrosoftWordE240F14_key = '\002\200\000}',
	MicrosoftWordE240F15_key = '\002\200\000~',
	MicrosoftWordE240F16_key = '\002\200\000\177',
	MicrosoftWordE240Scroll_lock_key = '\002\200\000\221',
	MicrosoftWordE240Semi_colon_key = '\002\200\000\272',
	MicrosoftWordE240Equals_key = '\002\200\000\273',
	MicrosoftWordE240Comma_key = '\002\200\000\274',
	MicrosoftWordE240Hyphen_key = '\002\200\000\275',
	MicrosoftWordE240Period_key = '\002\200\000\276',
	MicrosoftWordE240Slash_key = '\002\200\000\277',
	MicrosoftWordE240Back_single_quote_key = '\002\200\000\300',
	MicrosoftWordE240Open_square_brace_key = '\002\200\000\333',
	MicrosoftWordE240Back_slash_key = '\002\200\000\334',
	MicrosoftWordE240Close_square_brace_key = '\002\200\000\335',
	MicrosoftWordE240Single_quote_key = '\002\200\000\336'
};
typedef enum MicrosoftWordE240 MicrosoftWordE240;

enum MicrosoftWordE241 {
	MicrosoftWordE241Olelink = '\002\201\000\000',
	MicrosoftWordE241Oleembed = '\002\201\000\001',
	MicrosoftWordE241Olecontrol = '\002\201\000\002'
};
typedef enum MicrosoftWordE241 MicrosoftWordE241;

enum MicrosoftWordE242 {
	MicrosoftWordE242OleverbPrimary = '\002\202\000\000',
	MicrosoftWordE242OleverbShow = '\002\201\377\377',
	MicrosoftWordE242OleverbOpen = '\002\201\377\376',
	MicrosoftWordE242OleverbHide = '\002\201\377\375',
	MicrosoftWordE242OleverbUiactivate = '\002\201\377\374',
	MicrosoftWordE242OleverbInPlaceActivate = '\002\201\377\373',
	MicrosoftWordE242OleverbDiscardUndoState = '\002\201\377\372'
};
typedef enum MicrosoftWordE242 MicrosoftWordE242;

enum MicrosoftWordE243 {
	MicrosoftWordE243InLine = '\002\203\000\000',
	MicrosoftWordE243FloatOverText = '\002\203\000\001'
};
typedef enum MicrosoftWordE243 MicrosoftWordE243;

enum MicrosoftWordE244 {
	MicrosoftWordE244LeftPortrait = '\002\204\000\000',
	MicrosoftWordE244CenterPortrait = '\002\204\000\001',
	MicrosoftWordE244RightPortrait = '\002\204\000\002',
	MicrosoftWordE244LeftLandscape = '\002\204\000\003',
	MicrosoftWordE244CenterLandscape = '\002\204\000\004',
	MicrosoftWordE244RightLandscape = '\002\204\000\005',
	MicrosoftWordE244LeftClockwise = '\002\204\000\006',
	MicrosoftWordE244CenterClockwise = '\002\204\000\007',
	MicrosoftWordE244RightClockwise = '\002\204\000\010'
};
typedef enum MicrosoftWordE244 MicrosoftWordE244;

enum MicrosoftWordE245 {
	MicrosoftWordE245FullBlock = '\002\205\000\000',
	MicrosoftWordE245ModifiedBlock = '\002\205\000\001',
	MicrosoftWordE245SemiBlock = '\002\205\000\002'
};
typedef enum MicrosoftWordE245 MicrosoftWordE245;

enum MicrosoftWordE246 {
	MicrosoftWordE246LetterTop = '\002\206\000\000',
	MicrosoftWordE246LetterBottom = '\002\206\000\001',
	MicrosoftWordE246LetterLeft = '\002\206\000\002',
	MicrosoftWordE246LetterRight = '\002\206\000\003'
};
typedef enum MicrosoftWordE246 MicrosoftWordE246;

enum MicrosoftWordE247 {
	MicrosoftWordE247SalutationInformal = '\002\207\000\000',
	MicrosoftWordE247SalutationFormal = '\002\207\000\001',
	MicrosoftWordE247SalutationBusiness = '\002\207\000\002',
	MicrosoftWordE247SalutationOther = '\002\207\000\003'
};
typedef enum MicrosoftWordE247 MicrosoftWordE247;

enum MicrosoftWordE248 {
	MicrosoftWordE248GenderFemale = '\002\210\000\000',
	MicrosoftWordE248GenderMale = '\002\210\000\001',
	MicrosoftWordE248GenderNeutral = '\002\210\000\002',
	MicrosoftWordE248GenderUnknown = '\002\210\000\003'
};
typedef enum MicrosoftWordE248 MicrosoftWordE248;

enum MicrosoftWordE249 {
	MicrosoftWordE249ByMoving = '\002\211\000\000',
	MicrosoftWordE249BySelecting = '\002\211\000\001'
};
typedef enum MicrosoftWordE249 MicrosoftWordE249;

enum MicrosoftWordE250 {
	MicrosoftWordE250UndefinedConstant = '\000\230\226\177',
	MicrosoftWordE250ToggleConstant = '\000\230\226~',
	MicrosoftWordE250ForwardConstant = '\?\377\377\377',
	MicrosoftWordE250BackwardConstant = '\300\000\000\001',
	MicrosoftWordE250AutoPosition = '\000\000\000\000',
	MicrosoftWordE250FirstConstant = '\000\000\000\001',
	MicrosoftWordE250CreatorCode = 'MSWD'
};
typedef enum MicrosoftWordE250 MicrosoftWordE250;

enum MicrosoftWordE251 {
	MicrosoftWordE251PasteOleobject = '\002\213\000\000',
	MicrosoftWordE251PasteRtf = '\002\213\000\001',
	MicrosoftWordE251PasteText = '\002\213\000\002',
	MicrosoftWordE251PasteMetafilePicture = '\002\213\000\003',
	MicrosoftWordE251PasteBitmap = '\002\213\000\004',
	MicrosoftWordE251PasteDeviceIndependentBitmap = '\002\213\000\005',
	MicrosoftWordE251PasteHyperlink = '\002\213\000\007',
	MicrosoftWordE251PasteShape = '\002\213\000\010',
	MicrosoftWordE251PasteEnhancedMetafile = '\002\212\377\377',
	MicrosoftWordE251PasteStyledText = '\002\213\000\011',
	MicrosoftWordE251PasteHtml = '\002\213\000\012',
	MicrosoftWordE251PastePDF = '\002\213\000\013'
};
typedef enum MicrosoftWordE251 MicrosoftWordE251;

enum MicrosoftWordE252 {
	MicrosoftWordE252PrintDocumentContent = '\002\214\000\000',
	MicrosoftWordE252PrintProperties = '\002\214\000\001',
	MicrosoftWordE252PrintComments = '\002\214\000\002',
	MicrosoftWordE252PrintStyles = '\002\214\000\003',
	MicrosoftWordE252PrintAutoTextEntries = '\002\214\000\004',
	MicrosoftWordE252PrintKeyAssignments = '\002\214\000\005',
	MicrosoftWordE252PrintEnvelope = '\002\214\000\006'
};
typedef enum MicrosoftWordE252 MicrosoftWordE252;

enum MicrosoftWordE253 {
	MicrosoftWordE253PrintAllPages = '\002\215\000\000',
	MicrosoftWordE253PrintOddPagesOnly = '\002\215\000\001',
	MicrosoftWordE253PrintEvenPagesOnly = '\002\215\000\002'
};
typedef enum MicrosoftWordE253 MicrosoftWordE253;

enum MicrosoftWordE254 {
	MicrosoftWordE254PrintAllDocument = '\002\216\000\000',
	MicrosoftWordE254PrintSelection = '\002\216\000\001',
	MicrosoftWordE254PrintCurrentPage = '\002\216\000\002',
	MicrosoftWordE254PrintFromTo = '\002\216\000\003',
	MicrosoftWordE254PrintRangeOfPages = '\002\216\000\004'
};
typedef enum MicrosoftWordE254 MicrosoftWordE254;

enum MicrosoftWordE255 {
	MicrosoftWordE255Spelling = '\002\217\000\000',
	MicrosoftWordE255Grammar = '\002\217\000\001',
	MicrosoftWordE255Thesaurus = '\002\217\000\002',
	MicrosoftWordE255Hyphenation = '\002\217\000\003',
	MicrosoftWordE255SpellingComplete = '\002\217\000\004',
	MicrosoftWordE255SpellingCustom = '\002\217\000\005',
	MicrosoftWordE255SpellingLegal = '\002\217\000\006',
	MicrosoftWordE255SpellingMedical = '\002\217\000\007',
	MicrosoftWordE255HangulHanjaConversion = '\002\217\000\010',
	MicrosoftWordE255HangulHanjaConversionCustom = '\002\217\000\011'
};
typedef enum MicrosoftWordE255 MicrosoftWordE255;

enum MicrosoftWordE256 {
	MicrosoftWordE256SpellingWordTypeSpellWord = '\002\220\000\000',
	MicrosoftWordE256SpellingWordTypeWildcard = '\002\220\000\001',
	MicrosoftWordE256SpellingWordTypeAnagram = '\002\220\000\002'
};
typedef enum MicrosoftWordE256 MicrosoftWordE256;

enum MicrosoftWordE257 {
	MicrosoftWordE257SpellingCorrect = '\002\221\000\000',
	MicrosoftWordE257SpellingNotInDictionary = '\002\221\000\001',
	MicrosoftWordE257SpellingCapitalization = '\002\221\000\002'
};
typedef enum MicrosoftWordE257 MicrosoftWordE257;

enum MicrosoftWordE258 {
	MicrosoftWordE258SpellingError = '\002\222\000\000',
	MicrosoftWordE258GrammaticalError = '\002\222\000\001'
};
typedef enum MicrosoftWordE258 MicrosoftWordE258;

enum MicrosoftWordE259 {
	MicrosoftWordE259InlineShapeEmbeddedOleobject = '\002\223\000\001',
	MicrosoftWordE259InlineShapeLinkedOleobject = '\002\223\000\002',
	MicrosoftWordE259InlineShapePicture = '\002\223\000\003',
	MicrosoftWordE259InlineShapeLinkedPicture = '\002\223\000\004',
	MicrosoftWordE259InlineShapeOlecontrolObject = '\002\223\000\005',
	MicrosoftWordE259InlineShapeHorizontalLine = '\002\223\000\006',
	MicrosoftWordE259InlineShapePictureHorizontalLine = '\002\223\000\007',
	MicrosoftWordE259InlineShapeLinkedPictureHorizontalLine = '\002\223\000\010',
	MicrosoftWordE259InlineShapePictureBullet = '\002\223\000\011',
	MicrosoftWordE259InlineShapeScriptAnchor = '\002\223\000\012',
	MicrosoftWordE259InlineShapeOWSAnchor = '\002\223\000\013',
	MicrosoftWordE259InlineShapeChart = '\002\223\000\014',
	MicrosoftWordE259InlineShapeDiagram = '\002\223\000\015',
	MicrosoftWordE259InlineShapeLockedCanvas = '\002\223\000\016',
	MicrosoftWordE259InlineShapeSmartartGraphic = '\002\223\000\017'
};
typedef enum MicrosoftWordE259 MicrosoftWordE259;

enum MicrosoftWordE260 {
	MicrosoftWordE260Tiled = '\002\224\000\000',
	MicrosoftWordE260Icons = '\002\224\000\001'
};
typedef enum MicrosoftWordE260 MicrosoftWordE260;

enum MicrosoftWordE261 {
	MicrosoftWordE261SelectionStartActive = '\002\225\000\001',
	MicrosoftWordE261SelectionAtEol = '\002\225\000\002',
	MicrosoftWordE261SelectionOvertype = '\002\225\000\004',
	MicrosoftWordE261SelectionActive = '\002\225\000\010',
	MicrosoftWordE261SelectionReplace = '\002\225\000\020',
	MicrosoftWordE261SelectionInactive = '\002\225\000\000',
	MicrosoftWordE261SelectionStartActiveAndAtEol = '\002\225\000\003',
	MicrosoftWordE261SelectionStartActiveAndOvertype = '\002\225\000\005',
	MicrosoftWordE261SelectionAtEolAndOvertype = '\002\225\000\006',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndOvertype = '\002\225\000\007',
	MicrosoftWordE261SelectionStartActiveAndActive = '\002\225\000\011',
	MicrosoftWordE261SelectionAtEolAndActive = '\002\225\000\012',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndActive = '\002\225\000\013',
	MicrosoftWordE261SelectionOvertypeAndActive = '\002\225\000\014',
	MicrosoftWordE261SelectionStartActiveAndOvertypeAndActive = '\002\225\000\015',
	MicrosoftWordE261SelectionAtEolAndOvertypeAndActive = '\002\225\000\016',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndOvertypeAndActive = '\002\225\000\017',
	MicrosoftWordE261SelectionStartActiveAndReplace = '\002\225\000\021',
	MicrosoftWordE261SelectionAtEolAndReplace = '\002\225\000\022',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndReplace = '\002\225\000\023',
	MicrosoftWordE261SelectionOvertypeAndReplace = '\002\225\000\024',
	MicrosoftWordE261SelectionStartActiveAndOvertypeAndReplace = '\002\225\000\025',
	MicrosoftWordE261SelectionAtEolAndOvertypeAndReplace = '\002\225\000\026',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndOvertypeAndReplace = '\002\225\000\027',
	MicrosoftWordE261SelectionActiveAndReplace = '\002\225\000\030',
	MicrosoftWordE261SelectionStartActiveAndActiveAndReplace = '\002\225\000\031',
	MicrosoftWordE261SelectionAtEolAndActiveAndReplace = '\002\225\000\032',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndActiveAndReplace = '\002\225\000\033',
	MicrosoftWordE261SelectionOvertypeAndActiveAndReplace = '\002\225\000\034',
	MicrosoftWordE261SelectionStartActiveAndOvertypeAndActiveAndReplace = '\002\225\000\035',
	MicrosoftWordE261SelectionAtEolAndOvertypeAndActiveAndReplace = '\002\225\000\036',
	MicrosoftWordE261SelectionStartActiveAndAtEolAndOvertypeAndActiveAndReplace = '\002\225\000\037'
};
typedef enum MicrosoftWordE261 MicrosoftWordE261;

enum MicrosoftWordE262 {
	MicrosoftWordE262AutoVersionOff = '\002\226\000\000',
	MicrosoftWordE262AutoVersionOnClose = '\002\226\000\001'
};
typedef enum MicrosoftWordE262 MicrosoftWordE262;

enum MicrosoftWordE263 {
	MicrosoftWordE263OrganizerObjectStyles = '\002\227\000\000',
	MicrosoftWordE263OrganizerObjectAutoText = '\002\227\000\001',
	MicrosoftWordE263OrganizerObjectCommandBars = '\002\227\000\002'
};
typedef enum MicrosoftWordE263 MicrosoftWordE263;

enum MicrosoftWordE264 {
	MicrosoftWordE264MatchParagraphMark = '\002\231\000\017',
	MicrosoftWordE264MatchTabCharacter = '\002\230\000\011',
	MicrosoftWordE264MatchCommentMark = '\002\230\000\005',
	MicrosoftWordE264MatchAnyCharacter = '\002\231\000\?',
	MicrosoftWordE264MatchAnyDigit = '\002\231\000\037',
	MicrosoftWordE264MatchAnyLetter = '\002\231\000/',
	MicrosoftWordE264MatchCaretCharacter = '\002\230\000\013',
	MicrosoftWordE264MatchColumnBreak = '\002\230\000\016',
	MicrosoftWordE264MatchEmDash = '\002\230 \024',
	MicrosoftWordE264MatchEnDash = '\002\230 \023',
	MicrosoftWordE264MatchEndnoteMark = '\002\231\000\023',
	MicrosoftWordE264MatchField = '\002\230\000\023',
	MicrosoftWordE264MatchFootnoteMark = '\002\231\000\022',
	MicrosoftWordE264MatchGraphic = '\002\230\000\001',
	MicrosoftWordE264MatchManualLineBreak = '\002\231\000\017',
	MicrosoftWordE264MatchManualPageBreak = '\002\231\000\034',
	MicrosoftWordE264MatchNonbreakingHyphen = '\002\230\000\036',
	MicrosoftWordE264MatchNonbreakingSpace = '\002\230\000\240',
	MicrosoftWordE264MatchOptionalHyphen = '\002\230\000\037',
	MicrosoftWordE264MatchSectionBreak = '\002\231\000,',
	MicrosoftWordE264MatchWhiteSpace = '\002\231\000w'
};
typedef enum MicrosoftWordE264 MicrosoftWordE264;

enum MicrosoftWordE265 {
	MicrosoftWordE265FindStop = '\002\231\000\000',
	MicrosoftWordE265FindContinue = '\002\231\000\001',
	MicrosoftWordE265FindAsk = '\002\231\000\002'
};
typedef enum MicrosoftWordE265 MicrosoftWordE265;

enum MicrosoftWordE266 {
	MicrosoftWordE266ActiveEndAdjustedPageNumber = '\002\232\000\001',
	MicrosoftWordE266ActiveEndSectionNumber = '\002\232\000\002',
	MicrosoftWordE266ActiveEndPageNumber = '\002\232\000\003',
	MicrosoftWordE266NumberOfPagesInDocument = '\002\232\000\004',
	MicrosoftWordE266HorizontalPositionRelativeToPage = '\002\232\000\005',
	MicrosoftWordE266VerticalPositionRelativeToPage = '\002\232\000\006',
	MicrosoftWordE266HorizontalPositionRelativeToTextBoundary = '\002\232\000\007',
	MicrosoftWordE266VerticalPositionRelativeToTextBoundary = '\002\232\000\010',
	MicrosoftWordE266FirstCharacterColumnNumber = '\002\232\000\011',
	MicrosoftWordE266FirstCharacterLineNumber = '\002\232\000\012',
	MicrosoftWordE266FrameIsSelected = '\002\232\000\013',
	MicrosoftWordE266WithInTable = '\002\232\000\014',
	MicrosoftWordE266StartOfRangeRowNumber = '\002\232\000\015',
	MicrosoftWordE266End_ofRangeRowNumber = '\002\232\000\016',
	MicrosoftWordE266MaximumNumberOfRows = '\002\232\000\017',
	MicrosoftWordE266StartOfRangeColumnNumber = '\002\232\000\020',
	MicrosoftWordE266End_ofRangeColumnNumber = '\002\232\000\021',
	MicrosoftWordE266MaximumNumberOfColumns = '\002\232\000\022',
	MicrosoftWordE266ZoomPercentage = '\002\232\000\023',
	MicrosoftWordE266SelectionMode = '\002\232\000\024',
	MicrosoftWordE266InfoCapsLock = '\002\232\000\025',
	MicrosoftWordE266InfoNumLock = '\002\232\000\026',
	MicrosoftWordE266OverType = '\002\232\000\027',
	MicrosoftWordE266RevisionMarking = '\002\232\000\030',
	MicrosoftWordE266InFootnoteEndnotePane = '\002\232\000\031',
	MicrosoftWordE266InCommentPane = '\002\232\000\032',
	MicrosoftWordE266InHeaderFooter = '\002\232\000\034',
	MicrosoftWordE266AtEndOfRowMarker = '\002\232\000\037',
	MicrosoftWordE266ReferenceOfType = '\002\232\000 ',
	MicrosoftWordE266HeaderFooterType = '\002\232\000!',
	MicrosoftWordE266InMasterDocument = '\002\232\000\"',
	MicrosoftWordE266InFootnote = '\002\232\000#',
	MicrosoftWordE266InEndnote = '\002\232\000$',
	MicrosoftWordE266InWordMail = '\002\232\000%',
	MicrosoftWordE266InClipboard = '\002\232\000&'
};
typedef enum MicrosoftWordE266 MicrosoftWordE266;

enum MicrosoftWordE267 {
	MicrosoftWordE267WrapSquare = '\002\233\000\000',
	MicrosoftWordE267WrapTight = '\002\233\000\001',
	MicrosoftWordE267WrapThrough = '\002\233\000\002',
	MicrosoftWordE267WrapNone = '\002\233\000\003',
	MicrosoftWordE267WrapTopBottom = '\002\233\000\004'
};
typedef enum MicrosoftWordE267 MicrosoftWordE267;

enum MicrosoftWordE312 {
	MicrosoftWordE312PictureWrapTypeInlineWithText = '\002\310\000\000',
	MicrosoftWordE312PictureWrapTypeSquare = '\002\310\000\001',
	MicrosoftWordE312PictureWrapTypeTight = '\002\310\000\002',
	MicrosoftWordE312PictureWrapTypeBehindText = '\002\310\000\003',
	MicrosoftWordE312PictureWrapTypeInFrontOfText = '\002\310\000\004',
	MicrosoftWordE312PictureWrapTypeThrough = '\002\310\000\005',
	MicrosoftWordE312PictureWrapTypeTopAndBottom = '\002\310\000\006'
};
typedef enum MicrosoftWordE312 MicrosoftWordE312;

enum MicrosoftWordE268 {
	MicrosoftWordE268WrapBoth = '\002\234\000\000',
	MicrosoftWordE268WrapLeft = '\002\234\000\001',
	MicrosoftWordE268WrapRight = '\002\234\000\002',
	MicrosoftWordE268WrapLargest = '\002\234\000\003'
};
typedef enum MicrosoftWordE268 MicrosoftWordE268;

enum MicrosoftWordE269 {
	MicrosoftWordE269OutlineLevel1 = '\002\235\000\001',
	MicrosoftWordE269OutlineLevel2 = '\002\235\000\002',
	MicrosoftWordE269OutlineLevel3 = '\002\235\000\003',
	MicrosoftWordE269OutlineLevel4 = '\002\235\000\004',
	MicrosoftWordE269OutlineLevel5 = '\002\235\000\005',
	MicrosoftWordE269OutlineLevel6 = '\002\235\000\006',
	MicrosoftWordE269OutlineLevel7 = '\002\235\000\007',
	MicrosoftWordE269OutlineLevel8 = '\002\235\000\010',
	MicrosoftWordE269OutlineLevel9 = '\002\235\000\011',
	MicrosoftWordE269OutlineLevelBodyText = '\002\235\000\012'
};
typedef enum MicrosoftWordE269 MicrosoftWordE269;

enum MicrosoftWordE270 {
	MicrosoftWordE270TextOrientationHorizontal = '\002\236\000\000',
	MicrosoftWordE270TextOrientationUpward = '\002\236\000\002',
	MicrosoftWordE270TextOrientationDownward = '\002\236\000\003',
	MicrosoftWordE270TextOrientationVerticalEastAsian = '\002\236\000\001',
	MicrosoftWordE270TextOrientationHorizontalRotatedEastAsian = '\002\236\000\004'
};
typedef enum MicrosoftWordE270 MicrosoftWordE270;

enum MicrosoftWordE271 {
	MicrosoftWordE271ArtNone = '\002\237\000\000',
	MicrosoftWordE271ArtApples = '\002\237\000\001',
	MicrosoftWordE271ArtMapleMuffins = '\002\237\000\002',
	MicrosoftWordE271ArtCakeSlice = '\002\237\000\003',
	MicrosoftWordE271ArtCandyCorn = '\002\237\000\004',
	MicrosoftWordE271ArtIceCreamCones = '\002\237\000\005',
	MicrosoftWordE271ArtChampagneBottle = '\002\237\000\006',
	MicrosoftWordE271ArtPartyGlass = '\002\237\000\007',
	MicrosoftWordE271ArtChristmasTree = '\002\237\000\010',
	MicrosoftWordE271ArtTrees = '\002\237\000\011',
	MicrosoftWordE271ArtPalmsColor = '\002\237\000\012',
	MicrosoftWordE271ArtBalloons3Colors = '\002\237\000\013',
	MicrosoftWordE271ArtBalloonsHotAir = '\002\237\000\014',
	MicrosoftWordE271ArtPartyFavor = '\002\237\000\015',
	MicrosoftWordE271ArtConfettiStreamers = '\002\237\000\016',
	MicrosoftWordE271ArtHearts = '\002\237\000\017',
	MicrosoftWordE271ArtHeartBalloon = '\002\237\000\020',
	MicrosoftWordE271ArtStars3D = '\002\237\000\021',
	MicrosoftWordE271ArtStarsShadowed = '\002\237\000\022',
	MicrosoftWordE271ArtStars = '\002\237\000\023',
	MicrosoftWordE271ArtSun = '\002\237\000\024',
	MicrosoftWordE271ArtEarth2 = '\002\237\000\025',
	MicrosoftWordE271ArtEarth1 = '\002\237\000\026',
	MicrosoftWordE271ArtPeopleHats = '\002\237\000\027',
	MicrosoftWordE271ArtSombrero = '\002\237\000\030',
	MicrosoftWordE271ArtPencils = '\002\237\000\031',
	MicrosoftWordE271ArtPackages = '\002\237\000\032',
	MicrosoftWordE271ArtClocks = '\002\237\000\033',
	MicrosoftWordE271ArtFirecrackers = '\002\237\000\034',
	MicrosoftWordE271ArtRings = '\002\237\000\035',
	MicrosoftWordE271ArtMapPins = '\002\237\000\036',
	MicrosoftWordE271ArtConfetti = '\002\237\000\037',
	MicrosoftWordE271ArtCreaturesButterfly = '\002\237\000 ',
	MicrosoftWordE271ArtCreaturesLadyBug = '\002\237\000!',
	MicrosoftWordE271ArtCreaturesFish = '\002\237\000\"',
	MicrosoftWordE271ArtBirdsFlight = '\002\237\000#',
	MicrosoftWordE271ArtScaredCat = '\002\237\000$',
	MicrosoftWordE271ArtBats = '\002\237\000%',
	MicrosoftWordE271ArtFlowersRoses = '\002\237\000&',
	MicrosoftWordE271ArtFlowersRedRose = '\002\237\000\'',
	MicrosoftWordE271ArtPoinsettias = '\002\237\000(',
	MicrosoftWordE271ArtHolly = '\002\237\000)',
	MicrosoftWordE271ArtFlowersTiny = '\002\237\000*',
	MicrosoftWordE271ArtFlowersPansy = '\002\237\000+',
	MicrosoftWordE271ArtFlowersModern2 = '\002\237\000,',
	MicrosoftWordE271ArtFlowersModern1 = '\002\237\000-',
	MicrosoftWordE271ArtWhiteFlowers = '\002\237\000.',
	MicrosoftWordE271ArtVine = '\002\237\000/',
	MicrosoftWordE271ArtFlowersDaisies = '\002\237\0000',
	MicrosoftWordE271ArtFlowersBlockPrint = '\002\237\0001',
	MicrosoftWordE271ArtDecoArchColor = '\002\237\0002',
	MicrosoftWordE271ArtFans = '\002\237\0003',
	MicrosoftWordE271ArtFilm = '\002\237\0004',
	MicrosoftWordE271ArtLightning1 = '\002\237\0005',
	MicrosoftWordE271ArtCompass = '\002\237\0006',
	MicrosoftWordE271ArtDoubleD = '\002\237\0007',
	MicrosoftWordE271ArtClassicalWave = '\002\237\0008',
	MicrosoftWordE271ArtShadowedSquares = '\002\237\0009',
	MicrosoftWordE271ArtTwistedLines1 = '\002\237\000:',
	MicrosoftWordE271ArtWaveline = '\002\237\000;',
	MicrosoftWordE271ArtQuadrants = '\002\237\000<',
	MicrosoftWordE271ArtCheckedBarColor = '\002\237\000=',
	MicrosoftWordE271ArtSwirligig = '\002\237\000>',
	MicrosoftWordE271ArtPushPinNote1 = '\002\237\000\?',
	MicrosoftWordE271ArtPushPinNote2 = '\002\237\000@',
	MicrosoftWordE271ArtPumpkin1 = '\002\237\000A',
	MicrosoftWordE271ArtEggsBlack = '\002\237\000B',
	MicrosoftWordE271ArtCup = '\002\237\000C',
	MicrosoftWordE271ArtHeartGray = '\002\237\000D',
	MicrosoftWordE271ArtGingerbreadMan = '\002\237\000E',
	MicrosoftWordE271ArtBabyPacifier = '\002\237\000F',
	MicrosoftWordE271ArtBabyRattle = '\002\237\000G',
	MicrosoftWordE271ArtCabins = '\002\237\000H',
	MicrosoftWordE271ArtHouseFunky = '\002\237\000I',
	MicrosoftWordE271ArtStarsBlack = '\002\237\000J',
	MicrosoftWordE271ArtSnowflakes = '\002\237\000K',
	MicrosoftWordE271ArtSnowflakeFancy = '\002\237\000L',
	MicrosoftWordE271ArtSkyrocket = '\002\237\000M',
	MicrosoftWordE271ArtSeattle = '\002\237\000N',
	MicrosoftWordE271ArtMusicNotes = '\002\237\000O',
	MicrosoftWordE271ArtPalmsBlack = '\002\237\000P',
	MicrosoftWordE271ArtMapleLeaf = '\002\237\000Q',
	MicrosoftWordE271ArtPaperClips = '\002\237\000R',
	MicrosoftWordE271ArtShorebirdTracks = '\002\237\000S',
	MicrosoftWordE271ArtPeople = '\002\237\000T',
	MicrosoftWordE271ArtPeopleWaving = '\002\237\000U',
	MicrosoftWordE271ArtEclipsingSquares2 = '\002\237\000V',
	MicrosoftWordE271ArtHypnotic = '\002\237\000W',
	MicrosoftWordE271ArtDiamondsGray = '\002\237\000X',
	MicrosoftWordE271ArtDecoArch = '\002\237\000Y',
	MicrosoftWordE271ArtDecoBlocks = '\002\237\000Z',
	MicrosoftWordE271ArtCirclesLines = '\002\237\000[',
	MicrosoftWordE271ArtPapyrus = '\002\237\000\\',
	MicrosoftWordE271ArtWoodwork = '\002\237\000]',
	MicrosoftWordE271ArtWeavingBraid = '\002\237\000^',
	MicrosoftWordE271ArtWeavingRibbon = '\002\237\000_',
	MicrosoftWordE271ArtWeavingAngles = '\002\237\000`',
	MicrosoftWordE271ArtArchedScallops = '\002\237\000a',
	MicrosoftWordE271ArtSafari = '\002\237\000b',
	MicrosoftWordE271ArtCelticKnotwork = '\002\237\000c',
	MicrosoftWordE271ArtCrazyMaze = '\002\237\000d',
	MicrosoftWordE271ArtEclipsingSquares1 = '\002\237\000e',
	MicrosoftWordE271ArtBirds = '\002\237\000f',
	MicrosoftWordE271ArtFlowersTeacup = '\002\237\000g',
	MicrosoftWordE271ArtNorthwest = '\002\237\000h',
	MicrosoftWordE271ArtSouthwest = '\002\237\000i',
	MicrosoftWordE271ArtTribal6 = '\002\237\000j',
	MicrosoftWordE271ArtTribal4 = '\002\237\000k',
	MicrosoftWordE271ArtTribal3 = '\002\237\000l',
	MicrosoftWordE271ArtTribal2 = '\002\237\000m',
	MicrosoftWordE271ArtTribal5 = '\002\237\000n',
	MicrosoftWordE271ArtXillusions = '\002\237\000o',
	MicrosoftWordE271ArtZanyTriangles = '\002\237\000p',
	MicrosoftWordE271ArtPyramids = '\002\237\000q',
	MicrosoftWordE271ArtPyramidsAbove = '\002\237\000r',
	MicrosoftWordE271ArtConfettiGrays = '\002\237\000s',
	MicrosoftWordE271ArtConfettiOutline = '\002\237\000t',
	MicrosoftWordE271ArtConfettiWhite = '\002\237\000u',
	MicrosoftWordE271ArtMosaic = '\002\237\000v',
	MicrosoftWordE271ArtLightning2 = '\002\237\000w',
	MicrosoftWordE271ArtHeebieJeebies = '\002\237\000x',
	MicrosoftWordE271ArtLightBulb = '\002\237\000y',
	MicrosoftWordE271ArtGradient = '\002\237\000z',
	MicrosoftWordE271ArtTriangleParty = '\002\237\000{',
	MicrosoftWordE271ArtTwistedLines2 = '\002\237\000|',
	MicrosoftWordE271ArtMoons = '\002\237\000}',
	MicrosoftWordE271ArtOvals = '\002\237\000~',
	MicrosoftWordE271ArtDoubleDiamonds = '\002\237\000\177',
	MicrosoftWordE271ArtChainLink = '\002\237\000\200',
	MicrosoftWordE271ArtTriangles = '\002\237\000\201',
	MicrosoftWordE271ArtTribal1 = '\002\237\000\202',
	MicrosoftWordE271ArtMarqueeToothed = '\002\237\000\203',
	MicrosoftWordE271ArtSharksTeeth = '\002\237\000\204',
	MicrosoftWordE271ArtSawtooth = '\002\237\000\205',
	MicrosoftWordE271ArtSawtoothGray = '\002\237\000\206',
	MicrosoftWordE271ArtPostageStamp = '\002\237\000\207',
	MicrosoftWordE271ArtWeavingStrips = '\002\237\000\210',
	MicrosoftWordE271ArtZigZag = '\002\237\000\211',
	MicrosoftWordE271ArtCrossStitch = '\002\237\000\212',
	MicrosoftWordE271ArtGems = '\002\237\000\213',
	MicrosoftWordE271ArtCirclesRectangles = '\002\237\000\214',
	MicrosoftWordE271ArtCornerTriangles = '\002\237\000\215',
	MicrosoftWordE271ArtCreaturesInsects = '\002\237\000\216',
	MicrosoftWordE271ArtZigZagStitch = '\002\237\000\217',
	MicrosoftWordE271ArtCheckered = '\002\237\000\220',
	MicrosoftWordE271ArtCheckedBarBlack = '\002\237\000\221',
	MicrosoftWordE271ArtMarquee = '\002\237\000\222',
	MicrosoftWordE271ArtBasicWhiteDots = '\002\237\000\223',
	MicrosoftWordE271ArtBasicWideMidline = '\002\237\000\224',
	MicrosoftWordE271ArtBasicWideOutline = '\002\237\000\225',
	MicrosoftWordE271ArtBasicWideInline = '\002\237\000\226',
	MicrosoftWordE271ArtBasicThinLines = '\002\237\000\227',
	MicrosoftWordE271ArtBasicWhiteDashes = '\002\237\000\230',
	MicrosoftWordE271ArtBasicWhiteSquares = '\002\237\000\231',
	MicrosoftWordE271ArtBasicBlackSquares = '\002\237\000\232',
	MicrosoftWordE271ArtBasicBlackDashes = '\002\237\000\233',
	MicrosoftWordE271ArtBasicBlackDots = '\002\237\000\234',
	MicrosoftWordE271ArtStarsTop = '\002\237\000\235',
	MicrosoftWordE271ArtCertificateBanner = '\002\237\000\236',
	MicrosoftWordE271ArtHandmade1 = '\002\237\000\237',
	MicrosoftWordE271ArtHandmade2 = '\002\237\000\240',
	MicrosoftWordE271ArtTornPaper = '\002\237\000\241',
	MicrosoftWordE271ArtTornPaperBlack = '\002\237\000\242',
	MicrosoftWordE271ArtCouponCutoutDashes = '\002\237\000\243',
	MicrosoftWordE271ArtCouponCutoutDots = '\002\237\000\244'
};
typedef enum MicrosoftWordE271 MicrosoftWordE271;

enum MicrosoftWordE272 {
	MicrosoftWordE272BorderDistanceFromText = '\002\240\000\000',
	MicrosoftWordE272BorderDistanceFromPageEdge = '\002\240\000\001'
};
typedef enum MicrosoftWordE272 MicrosoftWordE272;

enum MicrosoftWordE273 {
	MicrosoftWordE273ReplaceNone = '\002\241\000\000',
	MicrosoftWordE273ReplaceOne = '\002\241\000\001',
	MicrosoftWordE273ReplaceAll = '\002\241\000\002'
};
typedef enum MicrosoftWordE273 MicrosoftWordE273;

enum MicrosoftWordE274 {
	MicrosoftWordE274FontBiasDoNotCare = '\002\242\000\377',
	MicrosoftWordE274FontBiasDefault = '\002\242\000\000',
	MicrosoftWordE274FontBiasEastAsian = '\002\242\000\001'
};
typedef enum MicrosoftWordE274 MicrosoftWordE274;

enum MicrosoftWordE275 {
	MicrosoftWordE275BrowserLevelV4 = '\002\243\000\000',
	MicrosoftWordE275BrowserLevelMicrosoftInternetExplorer5 = '\002\243\000\001'
};
typedef enum MicrosoftWordE275 MicrosoftWordE275;

enum MicrosoftWordE276 {
	MicrosoftWordE276EnclosureCircle = '\002\244\000\000',
	MicrosoftWordE276EnclosureSquare = '\002\244\000\001',
	MicrosoftWordE276EnclosureTriangle = '\002\244\000\002',
	MicrosoftWordE276EnclosureDiamond = '\002\244\000\003'
};
typedef enum MicrosoftWordE276 MicrosoftWordE276;

enum MicrosoftWordE277 {
	MicrosoftWordE277EncloseStyleNone = '\002\245\000\000',
	MicrosoftWordE277EncloseStyleSmall = '\002\245\000\001',
	MicrosoftWordE277EncloseStyleLarge = '\002\245\000\002'
};
typedef enum MicrosoftWordE277 MicrosoftWordE277;

enum MicrosoftWordE278 {
	MicrosoftWordE278LayoutModeDefault = '\002\246\000\000',
	MicrosoftWordE278LayoutModeGrid = '\002\246\000\001',
	MicrosoftWordE278LayoutModeLineGrid = '\002\246\000\002',
	MicrosoftWordE278LayoutModeGenko = '\002\246\000\003'
};
typedef enum MicrosoftWordE278 MicrosoftWordE278;

enum MicrosoftWordE279 {
	MicrosoftWordE279ForAEmailMessage = '\002\247\000\000',
	MicrosoftWordE279ForADocument = '\002\247\000\001',
	MicrosoftWordE279ForAWebPage = '\002\247\000\002'
};
typedef enum MicrosoftWordE279 MicrosoftWordE279;

enum MicrosoftWordE308 {
	MicrosoftWordE308TwoLinesInOneNone = '\002\304\000\000',
	MicrosoftWordE308TwoLinesInOneNoBrackets = '\002\304\000\001',
	MicrosoftWordE308TwoLinesInOneParentheses = '\002\304\000\002',
	MicrosoftWordE308TwoLinesInOneSquareBrackets = '\002\304\000\003',
	MicrosoftWordE308TwoLinesInOneAngleBrackets = '\002\304\000\004',
	MicrosoftWordE308TwoLinesInOneCurlyBrackets = '\002\304\000\005'
};
typedef enum MicrosoftWordE308 MicrosoftWordE308;

enum MicrosoftWordE309 {
	MicrosoftWordE309HorizontalInVerticalNone = '\002\305\000\000',
	MicrosoftWordE309HorizontalInVerticalFitInLine = '\002\305\000\001',
	MicrosoftWordE309HorizontalInVerticalResizeLine = '\002\305\000\002'
};
typedef enum MicrosoftWordE309 MicrosoftWordE309;

enum MicrosoftWordE280 {
	MicrosoftWordE280HorizontalLineAlignLeft = '\002\250\000\000',
	MicrosoftWordE280HorizontalLineAlignCenter = '\002\250\000\001',
	MicrosoftWordE280HorizontalLineAlignRight = '\002\250\000\002'
};
typedef enum MicrosoftWordE280 MicrosoftWordE280;

enum MicrosoftWordE281 {
	MicrosoftWordE281HorizontalLinePercentWidth = '\002\250\377\377',
	MicrosoftWordE281HorizontalLineFixedWidth = '\002\250\377\376'
};
typedef enum MicrosoftWordE281 MicrosoftWordE281;

enum MicrosoftWordE310 {
	MicrosoftWordE310PhoneticGuideAlignmentCenter = '\002\306\000\000',
	MicrosoftWordE310PhoneticGuideAlignmentZeroOneZero = '\002\306\000\001',
	MicrosoftWordE310PhoneticGuideAlignmentOneTwoOne = '\002\306\000\002',
	MicrosoftWordE310PhoneticGuideAlignmentLeft = '\002\306\000\003',
	MicrosoftWordE310PhoneticGuideAlignmentRight = '\002\306\000\004',
	MicrosoftWordE310PhoneticGuideAlignmentRightVertical = '\002\306\000\005'
};
typedef enum MicrosoftWordE310 MicrosoftWordE310;

enum MicrosoftWordE282 {
	MicrosoftWordE282TableDirectionRtl = '\002\252\000\000',
	MicrosoftWordE282TableDirectionLtr = '\002\252\000\001'
};
typedef enum MicrosoftWordE282 MicrosoftWordE282;

enum MicrosoftWordE283 {
	MicrosoftWordE283GutterPositionLeft = '\002\253\000\000',
	MicrosoftWordE283GutterPositionTop = '\002\253\000\001',
	MicrosoftWordE283GutterPositionRight = '\002\253\000\002'
};
typedef enum MicrosoftWordE283 MicrosoftWordE283;

enum MicrosoftWordE284 {
	MicrosoftWordE284GutterStyleLatin = '\002\253\377\366',
	MicrosoftWordE284GutterStyleBidi = '\002\254\000\002',
	MicrosoftWordE284GutterStyleNone = '\002\254\000\000'
};
typedef enum MicrosoftWordE284 MicrosoftWordE284;

enum MicrosoftWordE285 {
	MicrosoftWordE285ShapeTop = '\002\235\275\301',
	MicrosoftWordE285ShapeLeft = '\002\235\275\302',
	MicrosoftWordE285ShapeBottom = '\002\235\275\303',
	MicrosoftWordE285ShapeRight = '\002\235\275\304',
	MicrosoftWordE285ShapeCenter = '\002\235\275\305',
	MicrosoftWordE285ShapeInside = '\002\235\275\306',
	MicrosoftWordE285ShapeOutside = '\002\235\275\307'
};
typedef enum MicrosoftWordE285 MicrosoftWordE285;

enum MicrosoftWordE286 {
	MicrosoftWordE286TableTop = '\002\236\275\301',
	MicrosoftWordE286TableLeft = '\002\236\275\302',
	MicrosoftWordE286TableBottom = '\002\236\275\303',
	MicrosoftWordE286TableRight = '\002\236\275\304',
	MicrosoftWordE286TableCenter = '\002\236\275\305',
	MicrosoftWordE286TableInside = '\002\236\275\306',
	MicrosoftWordE286TableOutside = '\002\236\275\307'
};
typedef enum MicrosoftWordE286 MicrosoftWordE286;

enum MicrosoftWordE287 {
	MicrosoftWordE287Word8TableBehavior = '\002\257\000\000',
	MicrosoftWordE287Word9TableBehavior = '\002\257\000\001'
};
typedef enum MicrosoftWordE287 MicrosoftWordE287;

enum MicrosoftWordE288 {
	MicrosoftWordE288AutoFitFixed = '\002\260\000\000',
	MicrosoftWordE288AutoFitContent = '\002\260\000\001',
	MicrosoftWordE288AutoFitWindow = '\002\260\000\002'
};
typedef enum MicrosoftWordE288 MicrosoftWordE288;

enum MicrosoftWordE289 {
	MicrosoftWordE289Word8ListBehavior = '\002\261\000\000',
	MicrosoftWordE289Word9ListBehavior = '\002\261\000\001',
	MicrosoftWordE289Word10ListBehavior = '\002\261\000\002'
};
typedef enum MicrosoftWordE289 MicrosoftWordE289;

enum MicrosoftWordE290 {
	MicrosoftWordE290PreferredWidthAuto = '\002\262\000\001',
	MicrosoftWordE290PreferredWidthPercent = '\002\262\000\002',
	MicrosoftWordE290PreferredWidthPoints = '\002\262\000\003'
};
typedef enum MicrosoftWordE290 MicrosoftWordE290;

enum MicrosoftWordE291 {
	MicrosoftWordE291NewBlankDocument = '\002\263\000\000',
	MicrosoftWordE291NewWebPage = '\002\263\000\001',
	MicrosoftWordE291NewNotebookDocument = '\002\263\000\004',
	MicrosoftWordE291NewPublishingDocument = '\002\263\000\005'
};
typedef enum MicrosoftWordE291 MicrosoftWordE291;

enum MicrosoftWordE292 {
	MicrosoftWordE292UserFirst = '\002\264\000\001',
	MicrosoftWordE292UserLast = '\002\264\000\002',
	MicrosoftWordE292UserCompany = '\002\264\000\024',
	MicrosoftWordE292UserWorkStreet = '\002\264\000\026',
	MicrosoftWordE292UserWorkCity = '\002\264\000\027',
	MicrosoftWordE292UserWorkState = '\002\264\000\030',
	MicrosoftWordE292UserWorkZip = '\002\264\000\031',
	MicrosoftWordE292UserWorkPhone = '\002\264\000\035',
	MicrosoftWordE292UserEmailAddress1 = '\002\264\000f'
};
typedef enum MicrosoftWordE292 MicrosoftWordE292;

enum MicrosoftWordE293 {
	MicrosoftWordE293PasteDefault = '\002\265\000\000',
	MicrosoftWordE293SingleCellText = '\002\265\000\005',
	MicrosoftWordE293SingleCellTable = '\002\265\000\006',
	MicrosoftWordE293ListContinueNumbering = '\002\265\000\007',
	MicrosoftWordE293ListRestartNumbering = '\002\265\000\010',
	MicrosoftWordE293TableInsertAsRows = '\002\265\000\013',
	MicrosoftWordE293TableAppendTable = '\002\265\000\012',
	MicrosoftWordE293TableOriginalFormatting = '\002\265\000\014',
	MicrosoftWordE293ChartPicture = '\002\265\000\015',
	MicrosoftWordE293Chart = '\002\265\000\016',
	MicrosoftWordE293ChartLinked = '\002\265\000\017',
	MicrosoftWordE293FormatOriginalFormatting = '\002\265\000\020',
	MicrosoftWordE293FormatSurroundingFormattingWithEmphasis = '\002\265\000\024',
	MicrosoftWordE293FormatPlainText = '\002\265\000\026',
	MicrosoftWordE293TableOverwriteCells = '\002\265\000\027',
	MicrosoftWordE293ListCombineWithExistingList = '\002\265\000\030',
	MicrosoftWordE293ListDontMerge = '\002\265\000\031'
};
typedef enum MicrosoftWordE293 MicrosoftWordE293;

enum MicrosoftWordE294 {
	MicrosoftWordE294GoForward = '\002\266\000\001',
	MicrosoftWordE294GoBackward = '\002\266\000\002',
	MicrosoftWordE294ANumericConstant = '\002\266\000\000'
};
typedef enum MicrosoftWordE294 MicrosoftWordE294;

enum MicrosoftWordE311 {
	MicrosoftWordE311LineEndingCrLf = '\002\307\000\000',
	MicrosoftWordE311LineEndingCrOnly = '\002\307\000\001',
	MicrosoftWordE311LineEndingLfOnly = '\002\307\000\002',
	MicrosoftWordE311LineEndingLfCr = '\002\307\000\003',
	MicrosoftWordE311LineEndingLsPs = '\002\307\000\004'
};
typedef enum MicrosoftWordE311 MicrosoftWordE311;

enum MicrosoftWordE299 {
	MicrosoftWordE299ConditionFirstRow = '\002\301\000\000',
	MicrosoftWordE299ConditionLastRow = '\002\301\000\001',
	MicrosoftWordE299ConditionOddRowBanding = '\002\301\000\002',
	MicrosoftWordE299ConditionEvenRowBanding = '\002\301\000\003',
	MicrosoftWordE299ConditionFirstColumn = '\002\301\000\004',
	MicrosoftWordE299ConditionLastColumn = '\002\301\000\005',
	MicrosoftWordE299ConditionOddColumnBanding = '\002\301\000\006',
	MicrosoftWordE299ConditionEvenColumnBanding = '\002\301\000\007',
	MicrosoftWordE299ConditionTopRightCell = '\002\301\000\010',
	MicrosoftWordE299ConditionTopLeftCell = '\002\301\000\011',
	MicrosoftWordE299ConditionBottomRightCell = '\002\301\000\012',
	MicrosoftWordE299ConditionBottomLeftCell = '\002\301\000\013'
};
typedef enum MicrosoftWordE299 MicrosoftWordE299;

enum MicrosoftWordE295 {
	MicrosoftWordE295UnitALine = '\002\267\000\005',
	MicrosoftWordE295UnitAStory = '\002\267\000\006',
	MicrosoftWordE295UnitAScreen = '\002\267\000\007',
	MicrosoftWordE295UnitASection = '\002\267\000\010',
	MicrosoftWordE295UnitAColumn = '\002\267\000\011',
	MicrosoftWordE295UnitARow = '\002\267\000\012'
};
typedef enum MicrosoftWordE295 MicrosoftWordE295;

enum MicrosoftWordE296 {
	MicrosoftWordE296HighlightOn = '\002\270\377\377',
	MicrosoftWordE296HighlightOff = '\002\270\000\000',
	MicrosoftWordE296ANumericConstant = '\002\270\000\000'
};
typedef enum MicrosoftWordE296 MicrosoftWordE296;

enum MicrosoftWordE297 {
	MicrosoftWordE297CompareTargetSelected = '\002\271\000\000',
	MicrosoftWordE297CompareTargetCurrent = '\002\271\000\001',
	MicrosoftWordE297CompareTargetNew = '\002\271\000\002'
};
typedef enum MicrosoftWordE297 MicrosoftWordE297;

enum MicrosoftWordE298 {
	MicrosoftWordE298MergeTargetSelected = '\002\272\000\000',
	MicrosoftWordE298MergeTargetCurrent = '\002\272\000\001',
	MicrosoftWordE298MergeTargetNew = '\002\272\000\002'
};
typedef enum MicrosoftWordE298 MicrosoftWordE298;

enum MicrosoftWordE300 {
	MicrosoftWordE300RevisionsViewFinal = '\002\274\000\000',
	MicrosoftWordE300RevisionsViewOriginal = '\002\274\000\001'
};
typedef enum MicrosoftWordE300 MicrosoftWordE300;

enum MicrosoftWordE301 {
	MicrosoftWordE301RevisionsViewBalloons = '\002\275\000\000',
	MicrosoftWordE301RevisionsViewInline = '\002\275\000\001'
};
typedef enum MicrosoftWordE301 MicrosoftWordE301;

enum MicrosoftWordE303 {
	MicrosoftWordE303BalloonPrintOrientationAutomatic = '\002\277\000\000',
	MicrosoftWordE303BalloonPrintOrientationPreserve = '\002\277\000\001',
	MicrosoftWordE303BalloonPrintOrientationLandscape = '\002\277\000\002'
};
typedef enum MicrosoftWordE303 MicrosoftWordE303;

enum MicrosoftWordE304 {
	MicrosoftWordE304BalloonMarginLeft = '\002\300\000\000',
	MicrosoftWordE304BalloonMarginRight = '\002\300\000\001'
};
typedef enum MicrosoftWordE304 MicrosoftWordE304;

enum MicrosoftWordE331 {
	MicrosoftWordE331MinorVersion = '\002\333\000\000',
	MicrosoftWordE331MajorVersion = '\002\333\000\001',
	MicrosoftWordE331OverwriteCurrentVersion = '\002\333\000\002'
};
typedef enum MicrosoftWordE331 MicrosoftWordE331;

enum MicrosoftWordE332 {
	MicrosoftWordE332LockNone = '\002\334\000\000',
	MicrosoftWordE332LockReservation = '\002\334\000\001',
	MicrosoftWordE332LockEphemeral = '\002\334\000\002',
	MicrosoftWordE332LockChanged = '\002\334\000\003'
};
typedef enum MicrosoftWordE332 MicrosoftWordE332;

enum MicrosoftWord4017 {
	MicrosoftWord4017FieldOptions = 'w298',
	MicrosoftWord4017Field = 'w170'
};
typedef enum MicrosoftWord4017 MicrosoftWord4017;

enum MicrosoftWord4018 {
	MicrosoftWord4018FieldOptions = 'w298',
	MicrosoftWord4018Field = 'w170'
};
typedef enum MicrosoftWord4018 MicrosoftWord4018;

enum MicrosoftWord4024 {
	MicrosoftWord4024Revision = 'w219',
	MicrosoftWord4024Conflict = 'o120'
};
typedef enum MicrosoftWord4024 MicrosoftWord4024;

enum MicrosoftWord4025 {
	MicrosoftWord4025Revision = 'w219',
	MicrosoftWord4025Conflict = 'o120'
};
typedef enum MicrosoftWord4025 MicrosoftWord4025;

enum MicrosoftWord4013 {
	MicrosoftWord4013FootnoteOptions = 'WopN',
	MicrosoftWord4013EndnoteOptions = 'WopE'
};
typedef enum MicrosoftWord4013 MicrosoftWord4013;

enum MicrosoftWord4014 {
	MicrosoftWord4014FootnoteOptions = 'WopN',
	MicrosoftWord4014EndnoteOptions = 'WopE'
};
typedef enum MicrosoftWord4014 MicrosoftWord4014;

enum MicrosoftWord4015 {
	MicrosoftWord4015FootnoteOptions = 'WopN',
	MicrosoftWord4015EndnoteOptions = 'WopE'
};
typedef enum MicrosoftWord4015 MicrosoftWord4015;

enum MicrosoftWord4004 {
	MicrosoftWord4004Document = 'docu',
	MicrosoftWord4004Window = 'cwin',
	MicrosoftWord4004Pane = 'w120'
};
typedef enum MicrosoftWord4004 MicrosoftWord4004;

enum MicrosoftWord4019 {
	MicrosoftWord4019Font = 'w137',
	MicrosoftWord4019Frame = 'w175',
	MicrosoftWord4019SelectionObject = 'WSoj'
};
typedef enum MicrosoftWord4019 MicrosoftWord4019;

enum MicrosoftWord4021 {
	MicrosoftWord4021Field = 'w170',
	MicrosoftWord4021Frame = 'w175',
	MicrosoftWord4021FormField = 'w177',
	MicrosoftWord4021DataMergeField = 'w187',
	MicrosoftWord4021SelectionObject = 'WSoj',
	MicrosoftWord4021PageNumber = 'w225'
};
typedef enum MicrosoftWord4021 MicrosoftWord4021;

enum MicrosoftWord4023 {
	MicrosoftWord4023ListFormat = 'w123',
	MicrosoftWord4023WordList = 'w236'
};
typedef enum MicrosoftWord4023 MicrosoftWord4023;

enum MicrosoftWord4002 {
	MicrosoftWord4002Application = 'capp',
	MicrosoftWord4002Document = 'docu'
};
typedef enum MicrosoftWord4002 MicrosoftWord4002;

enum MicrosoftWord4003 {
	MicrosoftWord4003Application = 'capp',
	MicrosoftWord4003Document = 'docu'
};
typedef enum MicrosoftWord4003 MicrosoftWord4003;

enum MicrosoftWord4011 {
	MicrosoftWord4011Find = 'w124',
	MicrosoftWord4011Replacement = 'w125',
	MicrosoftWord4011SelectionObject = 'WSoj'
};
typedef enum MicrosoftWord4011 MicrosoftWord4011;

enum MicrosoftWord4012 {
	MicrosoftWord4012DropCap = 'w133',
	MicrosoftWord4012TabStop = 'w135',
	MicrosoftWord4012TextInput = 'w178',
	MicrosoftWord4012KeyBinding = 'w242'
};
typedef enum MicrosoftWord4012 MicrosoftWord4012;

enum MicrosoftWord4010 {
	MicrosoftWord4010Document = 'docu',
	MicrosoftWord4010ListFormat = 'w123',
	MicrosoftWord4010WordList = 'w236'
};
typedef enum MicrosoftWord4010 MicrosoftWord4010;

enum MicrosoftWord4020 {
	MicrosoftWord4020Field = 'w170',
	MicrosoftWord4020Frame = 'w175',
	MicrosoftWord4020FormField = 'w177',
	MicrosoftWord4020DataMergeField = 'w187',
	MicrosoftWord4020SelectionObject = 'WSoj',
	MicrosoftWord4020PageNumber = 'w225'
};
typedef enum MicrosoftWord4020 MicrosoftWord4020;

enum MicrosoftWord4005 {
	MicrosoftWord4005Window = 'cwin',
	MicrosoftWord4005Pane = 'w120'
};
typedef enum MicrosoftWord4005 MicrosoftWord4005;

enum MicrosoftWord4009 {
	MicrosoftWord4009Document = 'docu',
	MicrosoftWord4009ListFormat = 'w123',
	MicrosoftWord4009WordList = 'w236'
};
typedef enum MicrosoftWord4009 MicrosoftWord4009;

enum MicrosoftWord4007 {
	MicrosoftWord4007Window = 'cwin',
	MicrosoftWord4007Pane = 'w120'
};
typedef enum MicrosoftWord4007 MicrosoftWord4007;

enum MicrosoftWord4001 {
	MicrosoftWord4001Application = 'capp',
	MicrosoftWord4001Document = 'docu',
	MicrosoftWord4001Window = 'cwin'
};
typedef enum MicrosoftWord4001 MicrosoftWord4001;

enum MicrosoftWord4008 {
	MicrosoftWord4008Document = 'docu',
	MicrosoftWord4008ListFormat = 'w123',
	MicrosoftWord4008WordList = 'w236'
};
typedef enum MicrosoftWord4008 MicrosoftWord4008;

enum MicrosoftWord4006 {
	MicrosoftWord4006Window = 'cwin',
	MicrosoftWord4006Pane = 'w120'
};
typedef enum MicrosoftWord4006 MicrosoftWord4006;

enum MicrosoftWord4022 {
	MicrosoftWord4022TableOfFigures = 'w184',
	MicrosoftWord4022TableOfContents = 'w198'
};
typedef enum MicrosoftWord4022 MicrosoftWord4022;

enum MicrosoftWord4016 {
	MicrosoftWord4016LinkFormat = 'w167',
	MicrosoftWord4016FieldOptions = 'w298',
	MicrosoftWord4016TableOfFigures = 'w184',
	MicrosoftWord4016TableOfContents = 'w198',
	MicrosoftWord4016TableOfAuthorities = 'w200',
	MicrosoftWord4016Dialog = 'w202',
	MicrosoftWord4016Index = 'w215'
};
typedef enum MicrosoftWord4016 MicrosoftWord4016;

enum MicrosoftWordE315 {
	MicrosoftWordE315Def1 = '\002\312\377\377',
	MicrosoftWordE315Def2 = '\002\313\000\000',
	MicrosoftWordE315Def3 = '\002\313\000\001',
	MicrosoftWordE315Def4 = '\002\313\000\002',
	MicrosoftWordE315Def5 = '\002\313\000\003',
	MicrosoftWordE315Def6 = '\002\313\000\004',
	MicrosoftWordE315Def7 = '\002\313\000\005',
	MicrosoftWordE315Def8 = '\002\313\000\006',
	MicrosoftWordE315Def9 = '\002\313\000\007',
	MicrosoftWordE315Def10 = '\002\313\000\010',
	MicrosoftWordE315Def11 = '\002\313\000\011',
	MicrosoftWordE315Def12 = '\002\313\000\012',
	MicrosoftWordE315Def13 = '\002\313\000\013',
	MicrosoftWordE315Def14 = '\002\313\000\014',
	MicrosoftWordE315Def15 = '\002\313\000\015',
	MicrosoftWordE315Def16 = '\002\313\000\016',
	MicrosoftWordE315Def17 = '\002\313\000\017'
};
typedef enum MicrosoftWordE315 MicrosoftWordE315;

enum MicrosoftWord4027 {
	MicrosoftWord4027Shape = 'pShp',
	MicrosoftWord4027CalloutFormat = 'w275'
};
typedef enum MicrosoftWord4027 MicrosoftWord4027;

enum MicrosoftWord4028 {
	MicrosoftWord4028Callout = 'cD00',
	MicrosoftWord4028CalloutFormat = 'w275'
};
typedef enum MicrosoftWord4028 MicrosoftWord4028;

enum MicrosoftWord4029 {
	MicrosoftWord4029Callout = 'cD00',
	MicrosoftWord4029CalloutFormat = 'w275'
};
typedef enum MicrosoftWord4029 MicrosoftWord4029;

enum MicrosoftWord4030 {
	MicrosoftWord4030Callout = 'cD00',
	MicrosoftWord4030CalloutFormat = 'w275'
};
typedef enum MicrosoftWord4030 MicrosoftWord4030;

enum MicrosoftWord4031 {
	MicrosoftWord4031Shape = 'pShp',
	MicrosoftWord4031FillFormat = 'w278'
};
typedef enum MicrosoftWord4031 MicrosoftWord4031;

enum MicrosoftWord4032 {
	MicrosoftWord4032Shape = 'pShp',
	MicrosoftWord4032FillFormat = 'w278'
};
typedef enum MicrosoftWord4032 MicrosoftWord4032;

enum MicrosoftWord4033 {
	MicrosoftWord4033Shape = 'pShp',
	MicrosoftWord4033FillFormat = 'w278'
};
typedef enum MicrosoftWord4033 MicrosoftWord4033;

enum MicrosoftWord4034 {
	MicrosoftWord4034Shape = 'pShp',
	MicrosoftWord4034FillFormat = 'w278'
};
typedef enum MicrosoftWord4034 MicrosoftWord4034;

enum MicrosoftWord4035 {
	MicrosoftWord4035Shape = 'pShp',
	MicrosoftWord4035FillFormat = 'w278'
};
typedef enum MicrosoftWord4035 MicrosoftWord4035;

enum MicrosoftWord4036 {
	MicrosoftWord4036Shape = 'pShp',
	MicrosoftWord4036FillFormat = 'w278'
};
typedef enum MicrosoftWord4036 MicrosoftWord4036;

enum MicrosoftWord4037 {
	MicrosoftWord4037Shape = 'pShp',
	MicrosoftWord4037FillFormat = 'w278'
};
typedef enum MicrosoftWord4037 MicrosoftWord4037;

enum MicrosoftWord4038 {
	MicrosoftWord4038Shape = 'pShp',
	MicrosoftWord4038FillFormat = 'w278'
};
typedef enum MicrosoftWord4038 MicrosoftWord4038;

enum MicrosoftWord4039 {
	MicrosoftWord4039Shape = 'pShp',
	MicrosoftWord4039ThreeDFormat = 'w286'
};
typedef enum MicrosoftWord4039 MicrosoftWord4039;

enum MicrosoftWord4026 {
	MicrosoftWord4026Shape = 'pShp',
	MicrosoftWord4026InlineShape = 'w257'
};
typedef enum MicrosoftWord4026 MicrosoftWord4026;

enum MicrosoftWord4046 {
	MicrosoftWord4046Paragraph = 'cpar',
	MicrosoftWord4046ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4046 MicrosoftWord4046;

enum MicrosoftWord4040 {
	MicrosoftWord4040TextRange = 'w122',
	MicrosoftWord4040Section = 'w130',
	MicrosoftWord4040Paragraph = 'cpar',
	MicrosoftWord4040ParagraphFormat = 'w136',
	MicrosoftWord4040WordStyle = 'w173'
};
typedef enum MicrosoftWord4040 MicrosoftWord4040;

enum MicrosoftWord4041 {
	MicrosoftWord4041Paragraph = 'cpar',
	MicrosoftWord4041ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4041 MicrosoftWord4041;

enum MicrosoftWord4050 {
	MicrosoftWord4050Paragraph = 'cpar',
	MicrosoftWord4050ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4050 MicrosoftWord4050;

enum MicrosoftWord4051 {
	MicrosoftWord4051Paragraph = 'cpar',
	MicrosoftWord4051ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4051 MicrosoftWord4051;

enum MicrosoftWord4043 {
	MicrosoftWord4043Paragraph = 'cpar',
	MicrosoftWord4043ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4043 MicrosoftWord4043;

enum MicrosoftWord4042 {
	MicrosoftWord4042Paragraph = 'cpar',
	MicrosoftWord4042ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4042 MicrosoftWord4042;

enum MicrosoftWord4048 {
	MicrosoftWord4048Paragraph = 'cpar',
	MicrosoftWord4048ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4048 MicrosoftWord4048;

enum MicrosoftWord4047 {
	MicrosoftWord4047Paragraph = 'cpar',
	MicrosoftWord4047ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4047 MicrosoftWord4047;

enum MicrosoftWord4049 {
	MicrosoftWord4049Paragraph = 'cpar',
	MicrosoftWord4049ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4049 MicrosoftWord4049;

enum MicrosoftWord4044 {
	MicrosoftWord4044Paragraph = 'cpar',
	MicrosoftWord4044ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4044 MicrosoftWord4044;

enum MicrosoftWord4045 {
	MicrosoftWord4045Paragraph = 'cpar',
	MicrosoftWord4045ParagraphFormat = 'w136'
};
typedef enum MicrosoftWord4045 MicrosoftWord4045;

enum MicrosoftWord4058 {
	MicrosoftWord4058Row = 'crow',
	MicrosoftWord4058RowOptions = 'WrOp'
};
typedef enum MicrosoftWord4058 MicrosoftWord4058;

enum MicrosoftWord4059 {
	MicrosoftWord4059Row = 'crow',
	MicrosoftWord4059RowOptions = 'WrOp'
};
typedef enum MicrosoftWord4059 MicrosoftWord4059;

enum MicrosoftWord4060 {
	MicrosoftWord4060Column = 'ccol',
	MicrosoftWord4060ColumnOptions = 'WcOp'
};
typedef enum MicrosoftWord4060 MicrosoftWord4060;

enum MicrosoftWord4053 {
	MicrosoftWord4053Table = 'ctbl',
	MicrosoftWord4053Column = 'ccol'
};
typedef enum MicrosoftWord4053 MicrosoftWord4053;

enum MicrosoftWord4052 {
	MicrosoftWord4052Table = 'ctbl',
	MicrosoftWord4052Row = 'crow',
	MicrosoftWord4052Column = 'ccol',
	MicrosoftWord4052Cell = 'ccel',
	MicrosoftWord4052RowOptions = 'WrOp',
	MicrosoftWord4052ColumnOptions = 'WcOp',
	MicrosoftWord4052TableStyle = 'w292',
	MicrosoftWord4052Condition = 'w293'
};
typedef enum MicrosoftWord4052 MicrosoftWord4052;

enum MicrosoftWord4057 {
	MicrosoftWord4057Row = 'crow',
	MicrosoftWord4057Cell = 'ccel',
	MicrosoftWord4057RowOptions = 'WrOp'
};
typedef enum MicrosoftWord4057 MicrosoftWord4057;

enum MicrosoftWord4056 {
	MicrosoftWord4056Column = 'ccol',
	MicrosoftWord4056Cell = 'ccel',
	MicrosoftWord4056ColumnOptions = 'WcOp'
};
typedef enum MicrosoftWord4056 MicrosoftWord4056;

enum MicrosoftWord4054 {
	MicrosoftWord4054Table = 'ctbl',
	MicrosoftWord4054Column = 'ccol'
};
typedef enum MicrosoftWord4054 MicrosoftWord4054;

enum MicrosoftWord4055 {
	MicrosoftWord4055Table = 'ctbl',
	MicrosoftWord4055Column = 'ccol'
};
typedef enum MicrosoftWord4055 MicrosoftWord4055;



/*
 * Standard Suite
 */

// A scriptable object
@interface MicrosoftWordBaseObject : SBObject

@property (copy) NSDictionary *properties;  // All of the object's properties

- (void) closeSaving:(MicrosoftWordSavo)saving savingIn:(NSURL *)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) openFileName:(NSString *)fileName confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate Revert:(BOOL)Revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate fileConverter:(MicrosoftWordE162)fileConverter;  // Open the specified object(s). Returns a reference to the opened file when using the file name parameter.
- (void) printWithProperties:(MicrosoftWordPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(NSURL *)in_ as:(NSNumber *)as;  // Save an object
- (void) select;  // Make a selection
- (MicrosoftWordMathObject *) addMathArgumentBeforeArg:(NSInteger)beforeArg toFunction:(MicrosoftWordMathFunction *)toFunction toMatrixRow:(MicrosoftWordMathMatrixRow *)toMatrixRow toMatrixColumn:(MicrosoftWordMathMatrixColumn *)toMatrixColumn;  // Inserts an argument into an equation with variable number of arguments, i.e. math delimiter and math equation array objects, and returns a math object object. You must specify only one function, row, or argument.
- (void) autoPullLocksLocalDocument:(MicrosoftWordDocument *)localDocument serverDocument:(MicrosoftWordDocument *)serverDocument baseDocument:(MicrosoftWordDocument *)baseDocument;
- (void) autosaveDocument:(MicrosoftWordDocument *)document;  // Causes an autosave to occur for the given document or all documents if no document is specified.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Word can check out a specified document from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified document from a server to a local computer for editing. Returns a String that represents the local path and file name of the document checked out.
- (void) duplicatePage;  // Duplicates the page of the current selection and moves the selection to the new page. This command only works when in Publishing Layout View.
- (void) insertPage;  // Insert a page following the page of the current selection.
- (void) removePage;  // Removes the page of the selection and moves the selection to the following page. When removing the last page, the selection is moved to the previous page. This command only works when in Publishing Layout View.
- (void) threeWayMergeLocalDocument:(MicrosoftWordDocument *)localDocument serverDocument:(MicrosoftWordDocument *)serverDocument baseDocument:(MicrosoftWordDocument *)baseDocument favorSource:(BOOL)favorSource;
- (void) toggleShowCodes;

@end

// A basic application program
@interface MicrosoftWordBaseApplication : MicrosoftWordBaseObject

@property (readonly) BOOL frontmost;  // Is this the frontmost application?
@property (copy, readonly) NSString *name;  // the name
@property (copy, readonly) NSString *version;  // the version of the application


@end

// A basic document
@interface MicrosoftWordBaseDocument : MicrosoftWordBaseObject

@property (readonly) BOOL modified;  // Has the document been modified since the last save?
@property (copy, readonly) NSString *name;  // the name


@end

// A basic window
@interface MicrosoftWordBasicWindow : MicrosoftWordBaseObject

@property NSRect bounds;  // the boundary rectangle for the window
@property (readonly) BOOL closeable;  // Does the window have a close box?
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) NSInteger entryIndex;  // the number of the window
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL resizable;  // Is the window resizable?
@property (readonly) BOOL zoomable;  // Is the window zoomable?
@property BOOL zoomed;  // Is the window zoomed?
@property (copy, readonly) NSString *name;  // the title of the window
@property (readonly) BOOL visible;  // Is the window visible?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?


@end

@interface MicrosoftWordPrintSettings : SBObject

@property NSInteger copies;  // the number of copies of a document to be printed 
@property BOOL collating;  // Should printed copies be collated?
@property NSInteger startingPage;  // the first page of the document to be printed
@property NSInteger endingPage;  // the last page of the document to be printed
@property NSInteger pagesAcross;  // number of logical pages laid across a physical page
@property NSInteger pagesDown;  // number of logical pages laid out down a physical page
@property MicrosoftWordEnum errorHandling;  // how errors are handled
@property (copy) NSString *faxNumber;  // for fax number
@property (copy) NSString *targetPrinter;  // the queue name of the target printer

- (void) closeSaving:(MicrosoftWordSavo)saving savingIn:(NSURL *)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) openFileName:(NSString *)fileName confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate Revert:(BOOL)Revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate fileConverter:(MicrosoftWordE162)fileConverter;  // Open the specified object(s). Returns a reference to the opened file when using the file name parameter.
- (void) printWithProperties:(MicrosoftWordPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(NSURL *)in_ as:(NSNumber *)as;  // Save an object
- (void) select;  // Make a selection
- (MicrosoftWordMathObject *) addMathArgumentBeforeArg:(NSInteger)beforeArg toFunction:(MicrosoftWordMathFunction *)toFunction toMatrixRow:(MicrosoftWordMathMatrixRow *)toMatrixRow toMatrixColumn:(MicrosoftWordMathMatrixColumn *)toMatrixColumn;  // Inserts an argument into an equation with variable number of arguments, i.e. math delimiter and math equation array objects, and returns a math object object. You must specify only one function, row, or argument.
- (void) autoPullLocksLocalDocument:(MicrosoftWordDocument *)localDocument serverDocument:(MicrosoftWordDocument *)serverDocument baseDocument:(MicrosoftWordDocument *)baseDocument;
- (void) autosaveDocument:(MicrosoftWordDocument *)document;  // Causes an autosave to occur for the given document or all documents if no document is specified.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Word can check out a specified document from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified document from a server to a local computer for editing. Returns a String that represents the local path and file name of the document checked out.
- (void) duplicatePage;  // Duplicates the page of the current selection and moves the selection to the new page. This command only works when in Publishing Layout View.
- (void) insertPage;  // Insert a page following the page of the current selection.
- (void) removePage;  // Removes the page of the selection and moves the selection to the following page. When removing the last page, the selection is moved to the previous page. This command only works when in Publishing Layout View.
- (void) threeWayMergeLocalDocument:(MicrosoftWordDocument *)localDocument serverDocument:(MicrosoftWordDocument *)serverDocument baseDocument:(MicrosoftWordDocument *)baseDocument favorSource:(BOOL)favorSource;
- (void) toggleShowCodes;

@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface MicrosoftWordCommandBarControl : MicrosoftWordBaseObject

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) MicrosoftWordMCLT controlType;  // Returns the type of the command bar control.
@property (copy) NSString *descriptionText;  // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
@property BOOL enabled;  // Returns or sets if the command bar control is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number for this command bar control.
@property NSInteger height;  // Returns or sets the height of a command bar control.
@property NSInteger helpContextID;  // Returns or sets the help context ID number for the Help topic attached to the command bar control.
@property (copy) NSString *helpFile;  // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
- (NSInteger) id;  // Returns the id for a built-in command bar control.
@property (readonly) NSInteger leftPosition;  // Returns the left position of the command bar control.
@property (copy) NSString *name;  // Returns or sets the caption text for a command bar control.
@property (copy) NSString *parameter;  // Returns or sets a string that is used to execute a command.
@property NSInteger priority;  // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
@property (copy) NSString *tag;  // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
@property (copy) NSString *tooltipText;  // Returns or sets the text displayed in a command bar controls tooltip.
@property (readonly) NSInteger top;  // Returns the top position of a command bar control.
@property BOOL visible;  // Returns or sets if the command bar control is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar control.

- (void) execute;  // Runs the procedure or built-in command assigned to the specified command bar control.

@end

// A button control within a command bar.
@interface MicrosoftWordCommandBarButton : MicrosoftWordCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property MicrosoftWordMBns buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property MicrosoftWordMBTs buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface MicrosoftWordCommandBarCombobox : MicrosoftWordCommandBarControl

@property MicrosoftWordMXcb comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
@property (copy) NSString *comboboxText;  // Returns or sets the text in the display or edit portion of the command bar combobox control.
@property NSInteger dropDownLines;  // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
@property NSInteger dropDownWidth;  // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
@property NSInteger listIndex;

- (void) addItemToComboboxComboboxItem:(NSString *)comboboxItem entry_index:(NSInteger)entry_index;  // Add a new string to a custom combobox control.
- (void) clearCombobox;  // Clear all of the strings form a custom combobox.
- (NSString *) getComboboxItemEntry_index:(NSInteger)entry_index;  // Return the string at the given index within a combobox.
- (NSInteger) getCountOfComboboxItems;  // Return the number of strings within a combobox.
- (void) removeAnItemFromComboboxEntry_index:(NSInteger)entry_index;  // Remove a string from a custom combobox.
- (void) setComboboxItemEntry_index:(NSInteger)entry_index comboboxItem:(NSString *)comboboxItem;  // Set the string an a given index for a custom combobox.

@end

// A popup menu control within a command bar.
@interface MicrosoftWordCommandBarPopup : MicrosoftWordCommandBarControl

- (SBElementArray *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface MicrosoftWordCommandBar : MicrosoftWordBaseObject

- (SBElementArray *) commandBarControls;

@property MicrosoftWordMBPS barPosition;  // Returns or sets the position of the command bar.
@property (readonly) MicrosoftWordMBTY barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property (readonly) BOOL embeddable;  // Returns if the command bar can be displayed inside the document window.
@property BOOL embedded;  // Returns or sets if the command bar will be displayed inside the document window.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

@interface MicrosoftWordDocumentProperty : MicrosoftWordBaseObject

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface MicrosoftWordCustomDocumentProperty : MicrosoftWordDocumentProperty


@end

@interface MicrosoftWordWebPageFont : MicrosoftWordBaseObject

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft Word Suite
 */

// Represents a single comment.
@interface MicrosoftWordWordComment : MicrosoftWordBaseObject

@property (copy) NSString *author;
@property (readonly) NSInteger commentIndex;
@property (copy, readonly) MicrosoftWordTextRange *commentText;
@property (copy, readonly) NSDate *dateValue;
@property (copy) NSString *initials;
@property (copy, readonly) MicrosoftWordTextRange *noteReference;
@property (copy, readonly) MicrosoftWordTextRange *scope;
@property (readonly) BOOL showTip;


@end

// Represents a single list format that's been applied to specified paragraphs in a document.
@interface MicrosoftWordWordList : MicrosoftWordBaseObject

- (SBElementArray *) paragraphs;

@property (readonly) BOOL singleListTemplate;  // Returns if the entire Word list object uses the same list template.
@property (copy, readonly) NSString *styleName;
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the Word list.

- (void) applyListTemplateListTemplate:(MicrosoftWordListTemplate *)listTemplate continuePreviousList:(BOOL)continuePreviousList defaultListBehavior:(MicrosoftWordE289)defaultListBehavior;  // Applies a set of list-formatting characteristics to the specified Word list object.

@end

// Represents application and document options in Word.
@interface MicrosoftWordWordOptions : MicrosoftWordBaseObject

@property BOOL EnableMisusedWordsDictionary;  // Returns or sets if Microsoft Word checks for misused words when checking the spelling and grammar in a document.
@property BOOL IMEAutomaticControl;  // Returns or sets if Microsoft Word is set to automatically open and close the Japanese Input Method Editor. 
@property BOOL RTFInClipboard;  // Returns or sets if all text copied from Word to the Clipboard retains its character and paragraph formatting.
@property BOOL allowAccentedUppercase;  // Returns or sets if accents are retained when a French language character is changed to uppercase.
@property BOOL allowClickAndTypeMouse;  // Returns or sets if click and type functionality is enabled.
@property BOOL allowDragAndDrop;  // Returns or sets if dragging and dropping can be used to move or copy a selection.
@property BOOL allowFastSave;  // Returns or sets if Word saves only changes to a document. When reopening the document, Word uses the saved changes to reconstruct the document.
@property BOOL animateScreenMovements;  // Returns or sets if Word animates mouse movements, uses animated cursors, and animates actions such as background saving and find and replace operations.
@property BOOL applyEastAsianFontsToAscii;  // Returns or sets if Microsoft Word applies East Asian fonts to Latin text.
@property BOOL autoFormatApplyBulletedLists;  // Returns or sets if characters, such as asterisks, hyphens, and greater-than signs, at the beginning of list paragraphs are replaced with bullets from the Bullets and Numbering dialog box when Word formats a document or range automatically.
@property BOOL autoFormatApplyHeadings;  // Returns or sets if styles are automatically applied to headings when Word formats a document or range automatically.
@property BOOL autoFormatApplyLists;  // Returns or sets if  if styles are automatically applied to lists when Word formats a document or range automatically.
@property BOOL autoFormatApplyOtherParagraphs;  // Returns or sets if styles are automatically applied to paragraphs that aren't headings or list items when Word formats a document or range automatically.
@property BOOL autoFormatAsYouTypeApplyBorders;  // Returns or sets if a series of three or more hyphens -, equal signs =, or underscore characters _ are automatically replaced by a specific border line when the ENTER key is pressed.
@property BOOL autoFormatAsYouTypeApplyBulletedLists;  // Returns or sets if bullet characters, such as asterisks, hyphens, and greater-than signs, are replaced with bullets from the bullets and numbering dialog box as you type.
@property BOOL autoFormatAsYouTypeApplyClosings;  // Returns or sets if Microsoft Word to automatically applies the closing style to letter closings as you type.
@property BOOL autoFormatAsYouTypeApplyDates;  // Returns or sets if Microsoft Word  automatically applies the date style to dates as you type.
@property BOOL autoFormatAsYouTypeApplyFirstIndents;  // Returns or sets if Microsoft Word automatically replaces a space entered at the beginning of a paragraph with a first-line indent.
@property BOOL autoFormatAsYouTypeApplyHeadings;  // Returns or sets if styles are automatically applied to headings as you type.
@property BOOL autoFormatAsYouTypeApplyNumberedLists;  // Returns or sets if paragraphs are automatically formatted as numbered lists with a numbering scheme from the Bullets and Numbering dialog box according to what's typed.
@property BOOL autoFormatAsYouTypeApplyTables;  // Returns or set if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. The plus signs become the column borders, and the hyphens become the column widths.
@property BOOL autoFormatAsYouTypeAutoLetterWizard;  // Returns or sets if Microsoft Word to automatically starts the Letter Wizard when the user enters a letter salutation or closing.
@property BOOL autoFormatAsYouTypeDefineStyles;  // Returns or sets if Word automatically creates new styles based on manual formatting.
@property BOOL autoFormatAsYouTypeDeleteAutoSpaces;  // Returns or sets if Microsoft Word to automatically deletes spaces inserted between Japanese and Latin text as you type.
@property BOOL autoFormatAsYouTypeFormatListItemBeginning;  // Returns or sets if Word repeats character formatting applied to the beginning of a list item to the next list item.
@property BOOL autoFormatAsYouTypeInsertClosings;  // Returns or sets if Microsoft Word to automatically inserts the corresponding memo closing when the user enters a memo heading.
@property BOOL autoFormatAsYouTypeInsertOvers;  // Returns or sets if Word will automatically inset certain Japanese characters for other Japanese characters.
@property BOOL autoFormatAsYouTypeMatchParentheses;  // Returns or sets if Microsoft Word to automatically corrects improperly paired parentheses.
@property BOOL autoFormatAsYouTypeReplaceEastAsianDashes;  // Returns or sets if Microsoft Word to automatically corrects long vowel sounds and dashes.
@property BOOL autoFormatAsYouTypeReplaceFractions;  // Returns or sets if typed fractions are replaced with fractions from the current character set as you type.
@property BOOL autoFormatAsYouTypeReplaceHyperlinks;  // Returns or sets if e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLs, are automatically changed to hyperlinks as you type.
@property BOOL autoFormatAsYouTypeReplaceOrdinals;  // Returns or sets if the ordinal number suffixes st, nd, rd, and th are replaced with the same letters in superscript as you type. For example, 1st is replaced with 1 followed by st formatted as superscript.
@property BOOL autoFormatAsYouTypeReplacePlainTextEmphasis;  // Returns or sets if manual emphasis characters are automatically replaced with character formatting as you type.
@property BOOL autoFormatAsYouTypeReplaceQuotes;  // Returns or sets if straight quotation marks are automatically changed to smart, curly, quotation marks as you type.
@property BOOL autoFormatAsYouTypeReplaceSymbols;  // Returns or sets if two consecutive hyphens -- are replaced with an en dash – or an em dash — as you type.
@property BOOL autoFormatDeleteAutoSpaces;  // Returns or sets if Microsoft Word deletes spaces inserted between Japanese and Latin text, when Word formats a document or range automatically.
@property BOOL autoFormatMatchParentheses;  // Returns or sets if Microsoft Word  automatically corrects improperly paired parentheses.
@property BOOL autoFormatPreserveStyles;  // Returns or sets if previously applied styles are preserved when Word formats a document or range automatically.
@property BOOL autoFormatReplaceEastAsianDashes;  // Returns or sets if Microsoft Word automatically corrects long vowel sounds and dashes.
@property BOOL autoFormatReplaceFirstIndents;  // Returns or sets for Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent.
@property BOOL autoFormatReplaceFractions;  // Returns or sets if typed fractions are replaced with fractions from the current character set when Word formats a document or range automatically.
@property BOOL autoFormatReplaceHyperlinks;  // Returns or sets if e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLS, are changed to hyperlinks, when Word formats a document or range automatically.
@property BOOL autoFormatReplaceOrdinals;  // Returns or sets if the ordinal number suffixes st, nd, rd, and th are replaced with the same letters in superscript when Word formats a document or range automatically. For example, 1st is replaced with 1 followed by st formatted as superscript.
@property BOOL autoFormatReplacePlainTextEmphasis;  // Returns or sets if manual emphasis characters are replaced with character formatting when Word formats a document or range automatically.
@property BOOL autoFormatReplaceQuotes;  // Returns or sets if straight quotation marks are automatically changed to smart, curly, quotation marks when Word formats a document or range automatically.
@property BOOL autoFormatReplaceSymbols;  // Returns or set if two consecutive hyphens -- are replaced by an en dash – or an em dash — when Word formats a document or range automatically.
@property BOOL autoWordSelection;  // Returns or sets if dragging selects one word at a time instead of one character at a time.
@property BOOL automaticallyBuildUpEquations;  // Returns or sets whether Microsoft Word automatically converts equations to professional format.
@property BOOL ayMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL blueScreen;  // Returns or sets if Word displays text as white characters on a blue background.
@property NSInteger buttonFieldClicks;  // Returns or sets the number of clicks, either one or two, required to run a GOTOBUTTON or MACROBUTTON field.
@property BOOL bvMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL byteMatchFuzzy;  // Returns or sets Microsoft Word ignores the distinction between full-width and half-width characters, Latin or Japanese, during a search.
@property BOOL caseMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between uppercase and lowercase letters during a search.
@property BOOL checkGrammarAsYouType;  // Returns or sets if Word checks grammar and marks errors automatically as you type.
@property BOOL checkGrammarWithSpelling;  // Returns or sets if Word checks grammar while checking spelling.
@property BOOL checkSpellingAsYouType;  // Returns or sets if Word checks spelling and marks errors automatically as you type.
@property MicrosoftWordE110 commentsColor;  // Returns or sets the color of comments.
@property BOOL confirmConversions;  // Returns or sets if Word displays the convert file dialog box before it opens or inserts a file that isn't a Word document or template. In the convert file dialog box, the user chooses the format to convert the file from.
@property BOOL convertHighAnsiToEastAsian;  // Returns or sets if Microsoft Word converts text that is associated with an East Asian font to the appropriate font when it opens a document.
@property BOOL createBackup;  // Returns or sets if Word creates a backup copy each time a document is saved.
@property BOOL dashMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search.
@property (copy) NSColor *defaultBorderColor;  // Returns or sets the default RGB color to use for new border objects.
@property MicrosoftWordE110 defaultBorderColorIndex;  // Returns or sets the default line color index for borders.
@property MicrosoftWordMCoI defaultBorderColorThemeIndex;  // Returns or sets the default line color for borders.
@property MicrosoftWordE167 defaultBorderLineStyle;  // Returns or sets the default border line style
@property MicrosoftWordE168 defaultBorderLineWidth;  // Returns or sets the default line width of borders.
@property MicrosoftWordE110 defaultHighlightColorIndex;  // Returns or sets the color index  used to highlight text formatted with the highlight button.
@property MicrosoftWordE162 defaultOpenFormat;  // Returns or sets the default file converter used to open documents.
@property MicrosoftWordE318 deletedCellColor;  // Returns or sets the color of cells that are deleted while change tracking is enabled.
@property MicrosoftWordE110 deletedTextColor;  // Returns or sets the color of text that is deleted while change tracking is enabled.
@property MicrosoftWordE227 deletedTextMark;  // Returns or sets the format of text that is deleted while change tracking is enabled.
@property BOOL displayGridLines;  // Returns or sets if Microsoft Word displays the document grid. 
@property BOOL displayPasteOptions;  // Returns or sets if Microsoft Word  displays the Paste Options button, which displays directly under newly pasted text.
@property BOOL dzMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL enableSound;  // Returns or sets if Word makes the computer respond with a sound whenever an error occurs.
@property (readonly) BOOL envelopeFeederInstalled;  // Returns true if the current printer has a special feeder for envelopes.
@property BOOL fancyFontMenu;  // Returns or sets if the fancy font menu is shown.
@property double gridDistanceHorizontal;  // Returns or sets the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize autoshapes or East Asian characters in new documents.
@property double gridDistanceVertical;  // Returns or sets the amount of vertical space between the invisible gridlines that Word uses when you draw, move, and resize autoshapes or East Asian characters in new documents.
@property double gridOriginHorizontal;  // Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing autoshapes or East Asian characters to begin in new documents.
@property double gridOriginVertical;  // Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing autoshapes or East Asian characters to begin in new documents.
@property BOOL hfMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL hiraganaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between hiragana and katakana during a search.
@property BOOL ignoreInternetAndFileAddresses;  // Returns or sets if file name extensions,  paths, e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLs, are ignored while checking spelling.
@property BOOL ignoreMixedDigits;  // Returns or sets if words that contain numbers are ignored while checking spelling.
@property BOOL ignoreUppercase;  // Returns or sets if words in all uppercase letters are ignored while checking spelling.
@property BOOL inlineConversion;  // Returns or sets if Microsoft Word displays an unconfirmed character string in the Japanese Input Method Editor as an insertion between existing character strings.
@property BOOL insertKeyForPaste;  // Returns or sets if the insert key can be used for pasting the clipboard contents.
@property MicrosoftWordE318 insertedCellColor;  // Returns or sets the color of cells that are inserted while change tracking is enabled.
@property MicrosoftWordE110 insertedTextColor;  // Returns or sets the color of text that is inserted while change tracking is enabled.
@property MicrosoftWordE225 insertedTextMark;  // Returns or sets how Microsoft Word formats inserted text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the inserted text color property to control the look of inserted text.
@property BOOL iterationMarkMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between types of repetition marks during a search.
@property BOOL kanjiMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between standard and nonstandard kanji ideography during a search.
@property BOOL keepTrackOfFormatting;  // Returns or sets if Microsoft Word keeps track of formatting.
@property BOOL kiKuMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL liveWordCount;  // Returns or sets if the instant word count is displayed in the status bar.
@property BOOL mapPaperSize;  // Returns or sets if documents formatted for another country's or region's standard paper size, for example, A4 are automatically adjusted so that they're printed correctly on your country's/region's standard paper size, for example, Letter.
@property MicrosoftWordE171 measurementUnit;  // Returns or sets the standard measurement unit for Microsoft Word.
@property MicrosoftWordE318 mergedCellColor;  // Returns or sets the color of cells that are merged while change tracking is enabled.
@property MicrosoftWordE110 moveFromTextColor;  // Returns or sets the color of text that is moved from while change tracking is enabled.
@property MicrosoftWordE317 moveFromTextMark;  // Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text.
@property MicrosoftWordE110 moveToTextColor;  // Returns or sets the color of text that is moved to while change tracking is enabled.
@property MicrosoftWordE316 moveToTextMark;  // Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text.
@property BOOL oldKanaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between new kana and old kana characters during a search.
@property BOOL overtype;  // Returns or sets if overtype mode is active. In overtype mode, the characters you type replace existing characters one by one. When overtype isn't active, the characters you type move existing text to the right.
@property BOOL pagination;  // Returns or sets if Microsoft Word repaginates documents in the background.
@property BOOL pasteAdjustParagraphSpacing;  // Returns or sets if Microsoft Word automatically adjusts the spacing of paragraphs when cutting and pasting selections.
@property BOOL pasteAdjustTableFormatting;  // Returns or sets if Microsoft Word automatically adjusts the formatting of tables when cutting and pasting selections.
@property BOOL pasteAdjustWordSpacing;  // Returns or sets if Microsoft Word automatically adjusts the spacing of words when cutting and pasting selections.
@property BOOL pasteMergeFromExcel;  // Returns or sets if text formatting will be merged when pasting from Microsoft Excel.
@property BOOL pasteMergeFromPowerPoint;  // Returns or sets if text formatting will be merged when pasting from Microsoft PowerPoint.
@property BOOL pasteMergeLists;  // Returns or sets if the formatting of pasted lists will be merged with surrounding lists.
@property BOOL pasteSmartCutPaste;  // Returns or sets if Microsoft Word intelligently pastes selections into a document.
@property BOOL pasteSmartStyleBehavior;  // Returns or sets if Microsoft Word intelligently merges styles when pasting a selection from a different document.
@property (copy) NSString *pictureEditor;  // Returns or sets the name of the application to use to edit pictures.
@property MicrosoftWordE312 pictureWrapTypes;  // Returns or sets the wrapping that is used to insert or paste pictures.
@property BOOL plainEquationsUsePlainText;  // Returns or sets if equations are represented in plain text; false uses MathML
@property BOOL printComments;  // Returns or sets if Microsoft Word prints comments, starting on a new page at the end of the document.
@property BOOL printDrawingObjects;  // Returns or sets if Microsoft Word prints drawing objects.
@property BOOL printFieldCodes;  // Returns or sets if Microsoft Word prints field codes instead of field results.
@property BOOL printHiddenText;  // Returns or sets if hidden text is printed.
@property BOOL printProperties;  // Returns or sets if Microsoft Word prints document summary information on a separate page at the end of the document. 
@property BOOL printReverse;  // Returns or sets if Microsoft Word prints pages in reverse order.
@property BOOL prolongedSoundMarkMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between short and long vowel sounds during a search.
@property BOOL punctuationMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between types of punctuation marks during a search.
@property BOOL replaceSelection;  // Returns or sets if the result of typing or pasting replaces the selection. If false the result of typing or pasting is added before the selection, leaving the selection intact.
@property MicrosoftWordE110 revisedLinesColor;  // Returns or sets the color of changed lines in a document with tracked changes.
@property MicrosoftWordE226 revisedLinesMark;  // Returns or sets the placement of changed lines in a document with tracked changes.
@property MicrosoftWordE110 revisedPropertiesColor;  // Returns or sets the color index used to mark formatting changes while change tracking is enabled. 
@property MicrosoftWordE228 revisedPropertiesMark;  // Returns or sets the mark used to show formatting changes while change tracking is enabled.
@property NSInteger saveInterval;  // Returns or sets the time interval in minutes for saving autorecover information.
@property BOOL saveNormalPrompt;  // Returns or sets if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it quits. if this is set to false Word automatically saves changes to the Normal template before it quits.
@property BOOL savePropertiesPrompt;  // Returns or sets if Microsoft Word prompts for document property information when saving a new document.
@property BOOL sendMailAttach;  // True if the send to command on the file menu inserts the active document as an attachment to a mail message. False if the send to command inserts the contents of the active document as text in a mail message.
@property BOOL showReadabilityStatistics;  // Returns or sets if Microsoft Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar.
@property BOOL showWizardWelcome;  // Returns or sets if the welcome wizard should be shown.
@property BOOL smallKanaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between diphthongs and double consonants during a search.
@property BOOL smartCutPaste;  // Returns or sets if Microsoft Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs.
@property BOOL smartParagraphSelection;  // Returns or sets if Microsoft Word includes the paragraph mark in a selection when selecting most or all of a paragraph.
@property BOOL snapToGrid;  // Returns or sets if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in new documents.
@property BOOL snapToShapes;  // Returns or sets if Word automatically aligns autoshapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other autoshapes or East Asian characters in new documents.
@property BOOL spaceMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between space markers used during a search.
@property MicrosoftWordE318 splitCellColor;  // Returns or sets the color of cells that are split while change tracking is enabled.
@property BOOL suggestFromMainDictionaryOnly;  // Returns or sets if Microsoft Word draws spelling suggestions from the main dictionary only. If false it draws spelling suggestions from the main dictionary and any custom dictionaries that have been added.
@property BOOL suggestSpellingCorrections;  // Returns or sets if Microsoft Word always suggests alternative spellings for each misspelled word when checking spelling.
@property BOOL tabIndentKey;  // Returns or sets if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered and centered paragraphs to left-aligned.
@property BOOL tcMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL trackFormatting;  // Returns or sets tracking formatting flag.
@property BOOL trackMoves;  // Returns or sets tracking moves flag.
@property BOOL updateFieldsAtPrint;  // Returns or sets if Microsoft Word updates fields automatically before printing a document. 
@property BOOL updateLinksAtOpen;  // Returns or sets if Microsoft Word automatically updates all embedded OLE links in a document when it's opened.
@property BOOL updateLinksAtPrint;  // Returns or sets if Microsoft Word updates fields automatically before printing a document.
@property BOOL useCharacterUnit;  // Returns or sets if Microsoft Word uses characters as the default measurement unit for the current document.
@property BOOL useGermanSpellingReform;  // Returns or sets if Microsoft Word uses the German post-reform spelling rules when checking spelling.
@property BOOL warnBeforeSavingPrintingSendingMarkup;  // Returns or sets if Microsoft Word displays a warning when saving, printing, or sending as e-mail a document containing comments or tracked changes.
@property BOOL zjMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.


@end

// Represents a single add-in, either installed or not installed.
@interface MicrosoftWordAddIn : MicrosoftWordBaseObject

@property (readonly) BOOL autoload;  // Returns true if the specified add-in is automatically loaded when Word is started.
@property (readonly) BOOL compiled;  // Returns true if the specified add-in is a Word add-in library. False if the add-in is a template.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property BOOL installed;  // Returns or sets if the specified add-in is installed.
@property (copy, readonly) NSString *name;  // Returns the name of the add in.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified add-in.


@end

// An AppleScript object representing the Microsoft Word application.
@interface MicrosoftWordApplication : SBApplication

- (SBElementArray *) documents;
- (SBElementArray *) windows;
- (SBElementArray *) recentFiles;
- (SBElementArray *) fileConverters;
- (SBElementArray *) captionLabels;
- (SBElementArray *) addIns;
- (SBElementArray *) commandBars;
- (SBElementArray *) templates;
- (SBElementArray *) keyBindings;
- (SBElementArray *) dictionaries;

@property BOOL Word51Menus;  // Returns or sets if Word 5.1 menus should be used.
@property (copy, readonly) MicrosoftWordDocument *activeDocument;  // Returns the currently active document object.
@property (copy) NSString *activePrinter;  // Returns or sets the name of the active printer
@property (copy, readonly) MicrosoftWordWindow *activeWindow;  // Returns the currently active window object.
@property (copy, readonly) NSString *applicationVersion;  // Returns the Microsoft Word version number as a string.
@property (copy, readonly) MicrosoftWordAutocorrect *autocorrectObject;  // Returns the autocorrect object
@property (readonly) NSInteger backgroundPrintingStatus;  // Returns the number of print jobs in the background printing queue.
@property (copy) NSString *browseExtraFileTypes;  // Returns or sets to open hyperlinked HTML files in the Internet browser or Microsoft Word.  Set this property to text/html to allow hyperlinked HTML files to be opened in Microsoft Word.
@property (copy, readonly) MicrosoftWordBrowser *browserObject;  // Returns the browser object.  The browser object is a tool used to move the insertion point to objects in a document.
@property (copy, readonly) NSString *build;  // Returns the version and build number of the Microsoft Word application.
@property (readonly) BOOL capsLock;  // Returns if caps lock is turned on.
@property (copy, readonly) NSString *caption;  // Returns the name of the application.
@property (copy) SBObject *customizationContext;  // Returns or set a template or document object that represents the template or document in which changes to menus, toolbars, and key bindings are stored.
@property (copy) NSString *defaultSaveFormat;  // Returns or sets the default format. Common settings include: document = WordDocument, document template = Template, Word 97-2004 document = Doc97, Word XML document = XML, web page = Html, Text only = Text, RTF = Rtf, unicode text = Unicode.
@property (copy) NSString *defaultStartupPath;  // Returns or sets the complete path of the startup folder, excluding the final separator.
@property (copy) NSString *defaultTableSeparator;  // Returns or sets the single character used to separate text into cells when text is converted to a table.
@property (copy, readonly) MicrosoftWordDefaultWebOptions *defaultWebOptionsObject;  // Returns the default web options object.
@property MicrosoftWordE138 displayAlerts;  // Returns or sets the way certain alerts or messages are handled while an AppleScript is running.
@property BOOL displayAutoCompleteTips;  // Returns or set if Microsoft Word displays tips that suggest text for completing words, dates, or phrases as you type.
@property BOOL displayRecentFiles;  // Returns or sets if the names of recently used files are displayed on the file menu.
@property BOOL displayRibbon;  // Returns or sets if Word displays the Ribbon in at least one document window.  Setting this property to true will display the Ribbon in all windows.  Setting this property to false turns off all Ribbons in all windows.
@property BOOL displayScreenTips;  // Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted.
@property BOOL displayScrollBars;  // Returns or sets if Word displays a scroll bar in at least one document window.  Setting this property to true will display horizontal and vertical scrollbars in all windows.  Setting this property to false turns off all scrollbars in all windows.
@property BOOL displayStatusBar;  // Returns or sets if the status bar is displayed.
@property BOOL doPrintPreview;  // Returns or set if print preview is the current view.
@property (copy, readonly) NSArray *fontNames;  // Returns the list of names of all of the available fonts
@property NSInteger keyboardScript;  // Returns or sets the current keyboard script
@property (copy, readonly) NSArray *landscapeFontNames;  // Returns the list of names of all of the available landscape fonts
@property (readonly) NSInteger localizedLanguage;  // Returns the Windows language ID for the locale that Microsoft Word is using.
@property (copy, readonly) MicrosoftWordMailingLabel *mailingLabelObject;  // Returns the mailing label object.
@property (copy, readonly) MicrosoftWordMathAutocorrect *mathAc;  // Returns a math autocorrect object that represents the auto correct entries for equations.
@property (copy, readonly) NSString *name;  // Returns the name of this application.
@property (copy, readonly) MicrosoftWordTemplate *normalTemplate;  // Returns the normal template object
@property (readonly) BOOL numLock;  // Returns the state of the num lock key.  True if the keys on the numeric keypad insert numbers, false if the keys move the insertion point.
@property (copy, readonly) NSString *path;  // Returns the path to the application
@property (copy, readonly) NSString *pathSeparator;  // Returns the character used to separate folder names.
@property (copy, readonly) NSArray *portraitFontNames;  // Returns the list of names of all of the available portrait fonts
@property BOOL ribbonExpanded;  // Returns or sets if the Ribbon in expanded.
@property (copy, readonly) MicrosoftWordSelectionObject *selection;  // Returns the selection object.
@property (copy, readonly) MicrosoftWordWordOptions *settings;  // Returns the word options object.
@property BOOL showVisualBasicEditor;  // Return or set if the visual basic editor is visible.
@property (readonly) BOOL specialMode;  // Returns if Microsoft Word is in a special mode for example, copy text mode or move text mode.
@property BOOL startupDialog;  // Returns of sets if the project gallery dialog will be shown when starting Microsoft Word.
@property (copy) NSString *statusBar;  // Displays the specified text in the status bar. This is a write only property.
@property (copy, readonly) MicrosoftWordSystemObject *system_object;  // Returns the system object
@property (readonly) NSInteger usableHeight;  // Returns the maximum height in points to which you can set the width of a Microsoft Window document window.
@property (readonly) NSInteger usableWidth;  // Returns the maximum width in pixels to which you can set the width of a Microsoft Window document window.
@property (copy) NSString *userAddress;  // Returns of set the users mailing address.
@property (readonly) BOOL userControl;  // Returns true if the application was launched by a user.  False if the program was opened programmatically.
@property (copy) NSString *userInitials;  // Returns of sets the initials, which Microsoft Word uses to construct comment marks.
@property (copy) NSString *userName;  // Returns or sets the users name, which is used on envelopes and for the author document property.

- (void) quitSaving:(MicrosoftWordSavo)saving;  // Quit an application program
- (void) select;  // Make a selection
- (id) doScript:(NSString *)x;  // Execute a script
- (void) reset:(MicrosoftWord4000)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) WordHelpHelpType:(MicrosoftWordE238)helpType;  // Displays on-line Help information.
- (void) accept:(MicrosoftWord4024)x;  // Accepts the specified tracked change. The revision marks are removed, and the change is incorporated into the document.
- (void) activateObject:(MicrosoftWord4004)x;  // Activate this object.
- (MicrosoftWordAddIn *) addAddinFileName:(NSString *)fileName install:(BOOL)install;  // Returns an add in object that represents an add-in added to the list of available add-ins.
- (MicrosoftWordCoauthLock *) addCoauthLockToRange:(MicrosoftWordTextRange *)toRange toParagraph:(MicrosoftWordParagraph *)toParagraph toColumn:(MicrosoftWordColumn *)toColumn toCell:(MicrosoftWordCell *)toCell toRow:(MicrosoftWordRow *)toRow toTable:(MicrosoftWordTable *)toTable toSelection:(MicrosoftWordSelectionObject *)toSelection lockType:(MicrosoftWordE332)lockType inRange:(MicrosoftWordSelectionObject *)inRange inCoauthoring:(MicrosoftWordCoauthoring *)inCoauthoring inCoauthor:(MicrosoftWordCoauthor *)inCoauthor;  // Add a co-authoring lock.
- (void) arrangeWindowsArrangeStyle:(MicrosoftWordE260)arrangeStyle;  // Arrange windows on the screen.
- (NSInteger) buildKeyCodeKey1:(MicrosoftWordE240)key1 key2:(MicrosoftWordE240)key2 key3:(MicrosoftWordE240)key3 key4:(MicrosoftWordE240)key4;  // Returns a unique number for the specified key combination.
- (MicrosoftWordE102) canContinuePreviousList:(MicrosoftWord4023)x listTemplate:(MicrosoftWordListTemplate *)listTemplate;  // Returns whether the formatting from the previous list can be continued.
- (double) centimetersToPointsCentimeters:(double)centimeters;  // Converts a measurement from centimeters to points.
- (void) changeFileOpenDirectoryPath:(NSString *)path;  // Sets the folder in which Microsoft Word searches for documents.
- (BOOL) checkGrammar:(MicrosoftWord4002)x textToCheck:(NSString *)textToCheck;  // Checks a string for grammatical errors. Returns a boolean to indicate whether the string contains grammatical errors. True if the string contains no errors.
- (BOOL) checkSpelling:(MicrosoftWord4003)x textToCheck:(NSString *)textToCheck customDictionary:(MicrosoftWordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(MicrosoftWordDictionary *)mainDictionary customDictionary2:(MicrosoftWordDictionary *)customDictionary2 customDictionary3:(MicrosoftWordDictionary *)customDictionary3 customDictionary4:(MicrosoftWordDictionary *)customDictionary4 customDictionary5:(MicrosoftWordDictionary *)customDictionary5 customDictionary6:(MicrosoftWordDictionary *)customDictionary6 customDictionary7:(MicrosoftWordDictionary *)customDictionary7 customDictionary8:(MicrosoftWordDictionary *)customDictionary8 customDictionary9:(MicrosoftWordDictionary *)customDictionary9 customDictionary10:(MicrosoftWordDictionary *)customDictionary10;  // Checks a string for spelling errors. Returns true if the string has no spelling errors.
- (NSString *) cleanStringItemToCheck:(NSString *)itemToCheck;  // Removes nonprinting characters character codes 1 – 29 and special Microsoft Word characters from the specified string or changes them to spaces character code 32. Returns the result as a string.
- (void) clear:(MicrosoftWord4012)x;  // Removes the object.
- (void) clearFormatting:(MicrosoftWord4011)x;  // Removes text and paragraph formatting from a selection or from the formatting specified in a find or replace operation.
- (void) convertNumbersToText:(MicrosoftWord4009)x numberType:(MicrosoftWordE158)numberType;  // Changes the list numbers and listnum fields in the document object to text.
- (void) copyObject:(MicrosoftWord4020)x; //NS_RETURNS_NOT_RETAINED;  // Copies the content of the object to the clipboard.
- (NSInteger) countNumberedItems:(MicrosoftWord4010)x numberType:(MicrosoftWordE158)numberType level:(NSInteger)level;  // Returns the number of bulleted or numbered items and listnum fields in the document object.
- (MicrosoftWordDocument *) createNewDocumentAttachedTemplate:(MicrosoftWordTemplate *)attachedTemplate newTemplate:(BOOL)newTemplate newDocumentType:(MicrosoftWordE291)newDocumentType;  // Create a new document
- (void) createNewEquationFromRange:(MicrosoftWordTextRange *)fromRange inDocument:(MicrosoftWordDocument *)inDocument inRange:(MicrosoftWordSelectionObject *)inRange inSelection:(MicrosoftWordSelectionObject *)inSelection;  // Creates an equation, from the text equation contained within the specified range, and returns a text range object that contains the new equation.
- (void) createNewFieldTextRange:(MicrosoftWordTextRange *)textRange fieldType:(MicrosoftWordE183)fieldType fieldText:(NSString *)fieldText preserveFormatting:(BOOL)preserveFormatting;  // Create a new field
- (void) cutObject:(MicrosoftWord4021)x;  // Removes the object from the document and places it on the clipboard.
- (BOOL) doWordRepeatTimes:(NSInteger)times;  // Repeats the most recent editing action one or more times. Returns true if the commands were repeated successfully.
- (MicrosoftWordKeyBinding *) findKeyKeyCode:(NSInteger)keyCode key_code_2:(MicrosoftWordE240)key_code_2;  // Returns a key binding object that represents the specified key combination
- (MicrosoftWordBorder *) getBorder:(MicrosoftWord4019)x whichBorder:(MicrosoftWordE122)whichBorder;  // Returns the specified border object.
- (NSString *) getDefaultFilePathFilePathType:(MicrosoftWordE230)filePathType;  // Returns the default folders for items such as documents, templates, and graphics.
- (MicrosoftWordDialog *) getDialog:(MicrosoftWordE186)x;  // Returns a dialog object for the specified dialog.
- (NSString *) getInternationalInformation:(MicrosoftWordE115)x;  // Get the specified international information
- (NSArray *) getKeysBoundToKeyCategory:(MicrosoftWordE239)keyCategory command:(NSString *)command;  // Returns a list key binding objects that represents all the key combinations assigned to the specified item.
- (MicrosoftWordListGallery *) getListGallery:(MicrosoftWordE150)x;  // Returns the specified list gallery object object.
- (NSDictionary *) getSpellingSuggestionsItemToCheck:(NSString *)itemToCheck customDictionary:(MicrosoftWordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(MicrosoftWordDictionary *)mainDictionary suggestionMode:(MicrosoftWordE256)suggestionMode customDictionary2:(MicrosoftWordDictionary *)customDictionary2 customDictionary3:(MicrosoftWordDictionary *)customDictionary3 customDictionary4:(MicrosoftWordDictionary *)customDictionary4 customDictionary5:(MicrosoftWordDictionary *)customDictionary5 customDictionary6:(MicrosoftWordDictionary *)customDictionary6 customDictionary7:(MicrosoftWordDictionary *)customDictionary7 customDictionary8:(MicrosoftWordDictionary *)customDictionary8 customDictionary9:(MicrosoftWordDictionary *)customDictionary9 customDictionary10:(MicrosoftWordDictionary *)customDictionary10;  // Returns a record that specifies the error kind and a list of words.  The AEKeyword for the error kind is type and AEKeyword for the list is list.
- (MicrosoftWordSynonymInfo *) getSynonymInfoObjectItemToCheck:(NSString *)itemToCheck languageID:(MicrosoftWordE182)languageID;  // Returns a synonym info object that contains the information from the thesaurus on the synonyms, antonyms, or related words and expressions for the specified word or phrase.
- (NSString *) getUserPropertyPropertyType:(MicrosoftWordE292)propertyType;
- (MicrosoftWordWebPageFont *) getWebpageFont:(MicrosoftWordMChS)x;  // Return the specified web page font object for a given character set.
- (double) inchesToPointsInches:(double)inches;  // Converts a measurement from inches to points.
- (void) insertText:(NSString *)text at:(SBObject *)at;  // Insert the string at the specified location.
- (void) insertAutoTextAt:(MicrosoftWordTextRange *)at;  // Attempts to match the text in the specified range or the text surrounding the range with an existing auto text entry name.  If any such match is found, the auto text entry is inserted.  If no match an error occurs.
- (void) insertBreakAt:(MicrosoftWordTextRange *)at breakType:(MicrosoftWordE169)breakType;  // Inserts a break in the specified place of the specified kind.
- (void) insertCaptionAt:(MicrosoftWordTextRange *)at captionLabel:(MicrosoftWordE210)captionLabel title:(NSString *)title captionPosition:(MicrosoftWordE117)captionPosition;  // Inserts a caption immediately preceding or following the specified range.
- (void) insertCrossReferenceAt:(MicrosoftWordTextRange *)at referenceType:(MicrosoftWordE211)referenceType referenceKind:(MicrosoftWordE212)referenceKind referenceItem:(NSString *)referenceItem insertAsHyperlink:(BOOL)insertAsHyperlink includePosition:(BOOL)includePosition;  // Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined.
- (void) insertDatabaseAt:(MicrosoftWordTextRange *)at format:(MicrosoftWordE180)format style:(NSInteger)style linkToSource:(BOOL)linkToSource connection:(NSString *)connection SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1 passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate dataSource:(NSString *)dataSource from:(NSInteger)from to:(NSInteger)to includeFields:(BOOL)includeFields;  // Retrieves data from a data source and inserts the data as a table in place of the specified range.
- (void) insertDateTimeAt:(MicrosoftWordTextRange *)at dateTimeFormat:(NSString *)dateTimeFormat insertAsField:(BOOL)insertAsField;  // Insert the correct date or time, or both, either as text or as a time field at the specified location.
- (void) insertFileAt:(MicrosoftWordTextRange *)at fileName:(NSString *)fileName fileRange:(NSString *)fileRange confirmConversions:(BOOL)confirmConversions link:(BOOL)link;
- (void) insertParagraphAt:(MicrosoftWordTextRange *)at;  // Replaces the specified range with a new paragraph.  If you do not want to replace the range, use the collapse range method before using this method.
- (void) insertSymbolAt:(MicrosoftWordTextRange *)at characterNumber:(NSInteger)characterNumber font:(NSString *)font unicode:(BOOL)unicode bias:(MicrosoftWordE274)bias;  // Inserts a symbol in place of the specified range.  If you do not want to replace the range, use the collapse range method before using this method.
- (NSString *) keyStringKeyCode:(NSInteger)keyCode key_code_2:(MicrosoftWordE240)key_code_2;  // Returns the key combination string for the specified keys 
- (void) largeScroll:(MicrosoftWord4005)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls a window by the specified number of screens. This method is equivalent to clicking just before or just after the scroll boxes on the horizontal and vertical scroll bars.
- (double) linesToPointsLines:(double)lines;  // Converts a measurement from lines to points. 1 line = 12 points.
- (void) listCommandsListAllCommands:(BOOL)listAllCommands;  // Creates a new document and then inserts a table of Microsoft Word commands along with their associated shortcut keys and menu assignments.
- (double) millimetersToPointsMillimeters:(double)millimeters;  // Converts a measurement from millimeters to points.
- (void) organizerCopySource:(NSString *)source destination:(NSString *)destination name:(NSString *)name organizerObjectType:(MicrosoftWordE263)organizerObjectType;  // Copies the specified autotext entry, toolbar, style, or macro project item from the source document or template to the destination document or template.
- (void) organizerDeleteSource:(NSString *)source name:(NSString *)name organizerObjectType:(MicrosoftWordE263)organizerObjectType;  // Deletes the specified style, autotext entry, toolbar, or macro project item from a document or template.
- (void) organizerRenameSource:(NSString *)source name:(NSString *)name newName:(NSString *)newName organizerObjectType:(MicrosoftWordE263)organizerObjectType;  // Renames the specified style, autotext entry, toolbar, or macro project item in a document or template.
- (void) pageScroll:(MicrosoftWord4007)x down:(NSInteger)down up:(NSInteger)up;  // Scrolls through the window page by page.
- (double) picasToPointsPicas:(double)picas;  // Converts a measurement from picas to points.
- (double) pointsToCentimetersPoints:(double)points;  // Converts a measurement from points to centimeters.
- (double) pointsToInchesPoints:(double)points;  // Converts a measurement from points to inches.
- (double) pointsToLinesPoints:(double)points;  // Converts a measurement from points to lines. 1 line = 12 points.
- (double) pointsToMillimetersPoints:(double)points;  // Converts a measurement from points to millimeters.
- (double) pointsToPicasPoints:(double)points;  // Converts a measurement from points to picas.
- (void) printOut:(MicrosoftWord4001)x append:(BOOL)append printOutRange:(MicrosoftWordE254)printOutRange outputFileName:(NSString *)outputFileName pageFrom:(NSInteger)pageFrom pageTo:(NSInteger)pageTo printOutItem:(MicrosoftWordE252)printOutItem printCopies:(NSInteger)printCopies printOutPageType:(MicrosoftWordE253)printOutPageType printToFile:(BOOL)printToFile collate:(BOOL)collate fileName:(NSString *)fileName manualDuplexPrint:(BOOL)manualDuplexPrint;  // Prints out all or part of the specified or active document. This command has been deprecated; use the Print command in the Standard Suite.
- (void) reject:(MicrosoftWord4025)x;  // Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.
- (void) removeNumbers:(MicrosoftWord4008)x numberType:(MicrosoftWordE158)numberType;  // Removes numbers or bullets from the document
- (void) remveEmphemeralLocksInRange:(MicrosoftWordSelectionObject *)inRange inCoauthoring:(MicrosoftWordCoauthoring *)inCoauthoring inCoauthor:(MicrosoftWordCoauthor *)inCoauthor;
- (void) resetContinuationNotice:(MicrosoftWord4015)x;  // Resets the footnote or endnote continuation notice to the default notice. The default notice is blank.
- (void) resetContinuationSeparator:(MicrosoftWord4014)x;  // Resets the footnote or endnote continuation separator to the default separator. The default separator is a long horizontal line that separates document text from notes continued from the previous page.
- (void) resetIgnoreAll;  // Clears the list of words that were previously ignored during a spelling check. After you run this method, previously ignored words are checked along with all the other words.
- (void) resetSeparator:(MicrosoftWord4013)x;  // Resets the footnote or endnote separator to the default separator. The default separator is a short horizontal line that separates document text from notes.
- (MicrosoftWordLanguage *) retrieveLanguage:(MicrosoftWordE182)x;  // Returns the language object for the specified language.
- (void) runVBMacroMacroName:(NSString *)macroName;  // Runs a Visual Basic macro.
- (void) screenRefresh;  // Updates the display on the monitor with the current information in the video memory buffer. You can use this method after using the screen updating property to disable screen updates.
- (void) setDefaultFilePathFilePathType:(MicrosoftWordE230)filePathType path:(NSString *)path;  // Sets the default folders for items such as documents, templates, and graphics.
- (void) setUserPropertyPropertyType:(MicrosoftWordE292)propertyType propertyValue:(NSString *)propertyValue;
- (void) smallScroll:(MicrosoftWord4006)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls a window by the specified number of lines. This method is equivalent to clicking the scroll arrows on the horizontal and vertical scroll bars.
- (void) substituteFontUnavailableFont:(NSString *)unavailableFont substituteFont:(NSString *)substituteFont;  // Sets font-mapping options, which are reflected in the font substitution dialog box
- (void) unlink:(MicrosoftWord4017)x;
- (void) unloadAddinsRemoveFromList:(BOOL)removeFromList;  // Unloads all loaded add-ins and, depending on the value of the remove from list argument, removes them from the list of add-ins.
- (void) update:(MicrosoftWord4016)x;  // Updates the values shown in a built-in Microsoft Word dialog box, updates the specified link, or updates the entries shown in specified index, table of authorities, table of figures or table of contents.
- (void) updatePageNumbers:(MicrosoftWord4022)x;  // Updates the page numbers for items in the specified table of contents or table of figures.
- (void) updateSource:(MicrosoftWord4018)x;
- (void) automaticLength:(MicrosoftWord4027)x;  // Specifies that the first segment of the callout line the segment attached to the text callout box be scaled automatically when the callout is moved. Applies only to callouts whose lines consist of more than one segment.
- (void) customDrop:(MicrosoftWord4028)x drop:(double)drop;  // Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
- (void) customLength:(MicrosoftWord4029)x length:(double)length;  // Specifies that the first segment of the callout, line the segment attached to the text callout box, retain a fixed length whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment.
- (void) oneColorGradient:(MicrosoftWord4031)x gradientStyle:(MicrosoftWordMGdS)gradientStyle gradientVariant:(NSInteger)gradientVariant gradientDegree:(double)gradientDegree;  // Sets the specified fill to a one-color gradient.
- (void) patterned:(MicrosoftWord4032)x pattern:(MicrosoftWordPpTy)pattern;  // Sets the specified fill to a pattern.
- (void) presetDrop:(MicrosoftWord4030)x DropType:(MicrosoftWordMCDt)DropType;  // Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
- (void) presetGradient:(MicrosoftWord4033)x style:(MicrosoftWordMGdS)style gradientVariant:(NSInteger)gradientVariant presetGradientType:(MicrosoftWordMPGb)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetTextured:(MicrosoftWord4034)x presetTexture:(MicrosoftWordMPzT)presetTexture;  // Sets the specified fill to a preset texture
- (void) resetRotation:(MicrosoftWord4039)x;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- (void) saveAsPicture:(MicrosoftWord4026)x pictureType:(MicrosoftWordMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) solid:(MicrosoftWord4035)x;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) twoColorGradient:(MicrosoftWord4036)x gradientStyle:(MicrosoftWordMGdS)gradientStyle gradientVariant:(NSInteger)gradientVariant;  // Sets the specified fill to a two-color gradient.
- (void) userPicture:(MicrosoftWord4037)x pictureFile:(NSString *)pictureFile;  // Fills the specified shape with one large image. If you want to fill the shape with small tiles of an image, use the user textured method.
- (void) userTextured:(MicrosoftWord4038)x textureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the user picture method.
- (void) closeUp:(MicrosoftWord4041)x;  // Removes any spacing before the specified paragraphs.
- (MicrosoftWordBorder *) getBorderi:(MicrosoftWord4040)x whichBorder:(MicrosoftWordE122)whichBorder;  // Returns the specified border object.
- (void) indentCharWidth:(MicrosoftWord4050)x count:(NSInteger)count;  // Indents one or more paragraphs by a specified number of characters.
- (void) indentFirstLineCharWidth:(MicrosoftWord4051)x count:(NSInteger)count;  // Indents the first line of one or more paragraphs by a specified number of characters.
- (void) openOrCloseUp:(MicrosoftWord4043)x;  // If spacing before the specified paragraphs is zero, this method sets spacing to 12 points. If spacing before the paragraphs is greater than zero, this method sets spacing to zero.
- (void) openUp:(MicrosoftWord4042)x;  // Sets spacing before the specified paragraphs to 12 points.
- (void) reseti:(MicrosoftWord4046)x;  // Removes paragraph formatting that differs from the underlying style. For example, if you manually right align a paragraph and the underlying style has a different alignment, the reset method changes the alignment to match the style formatting.
- (void) space1:(MicrosoftWord4047)x;  // Single-spaces the specified paragraphs. The exact spacing is determined by the font size of the largest characters in each paragraph.
- (void) space15:(MicrosoftWord4048)x;  // Formats the specified paragraphs with 1.5-line spacing. The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.
- (void) space2:(MicrosoftWord4049)x;  // Double-spaces the specified paragraphs. The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.
- (void) tabHangingIndent:(MicrosoftWord4044)x count:(NSInteger)count;  // Sets a hanging indent to a specified number of tab stops. Can be used to remove tab stops from a hanging indent if the value of the count argument is a negative number.
- (void) tabIndent:(MicrosoftWord4045)x count:(NSInteger)count;  // Sets the left indent for the specified paragraphs to a specified number of tab stops. Can also be used to remove the indent if the value of the count argument is a negative number.
- (void) autoFit:(MicrosoftWord4060)x;  // Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
- (MicrosoftWordTextRange *) convertRowToText:(MicrosoftWord4059)x separator:(MicrosoftWordE177)separator nestedTables:(BOOL)nestedTables;  // Converts a row to text and returns a text range object that represents the delimited text.
- (MicrosoftWordBorder *) getBorderi2:(MicrosoftWord4052)x whichBorder:(MicrosoftWordE122)whichBorder;  // Returns the specified border object.
- (void) setLeftIndent:(MicrosoftWord4058)x leftIndent:(double)leftIndent rulerStyle:(MicrosoftWordE141)rulerStyle;  // Sets the indentation for a row or rows in a table.
- (void) setTableItemHeight:(MicrosoftWord4057)x rowHeight:(NSInteger)rowHeight heightRule:(MicrosoftWordE133)heightRule;  // Sets the height of table rows.
- (void) setTableItemWidth:(MicrosoftWord4056)x columnWidth:(double)columnWidth rulerStyle:(MicrosoftWordE141)rulerStyle;  // Sets the width of columns or cells in a table
- (void) sortAscending:(MicrosoftWord4054)x;  // Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) sortDescending:(MicrosoftWord4055)x;  // Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) tableSort:(MicrosoftWord4053)x excludeHeader:(BOOL)excludeHeader fieldNumber:(NSInteger)fieldNumber sortFieldType:(MicrosoftWordE178)sortFieldType sortOrder:(MicrosoftWordE179)sortOrder fieldNumberTwo:(NSInteger)fieldNumberTwo sortFieldTypeTwo:(MicrosoftWordE178)sortFieldTypeTwo sortOrderTwo:(MicrosoftWordE179)sortOrderTwo fieldNumberThree:(NSInteger)fieldNumberThree sortFieldTypeThree:(MicrosoftWordE178)sortFieldTypeThree sortOrderThree:(MicrosoftWordE179)sortOrderThree sortColumn:(BOOL)sortColumn separator:(MicrosoftWordE176)separator caseSensitive:(BOOL)caseSensitive languageID:(MicrosoftWordE182)languageID;  // Sort a table object.  For the column object only the first field is used

@end

// Represents a single AutoText entry.
@interface MicrosoftWordAutoTextEntry : MicrosoftWordBaseObject

@property (copy) NSString *autoTextValue;  // Returns or sets the value of this auto text entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of this auto text entry.
@property (copy, readonly) NSString *styleName;  // Returns the name of the style applied to the specified auto text entry.

- (MicrosoftWordTextRange *) insertAutoTextEntryWhere:(MicrosoftWordTextRange *)where richText:(BOOL)richText;  // Inserts the auto text entry in place of the specified range. Returns a text range object that represents the auto text entry.

@end

// Represents a single bookmark.
@interface MicrosoftWordBookmark : MicrosoftWordBaseObject

@property (readonly) BOOL column;  // True if the specified bookmark is a table column.
@property (readonly) BOOL empty;  // True if the specified bookmark is empty. An empty bookmark marks a location of a collapsed selection, it doesn't mark any text.
@property NSInteger endOfBookmark;  // Returns or sets the ending character position of the bookmark.
@property (copy, readonly) NSString *name;  // The name of the bookmark object.
@property NSInteger startOfBookmark;  // Returns or sets the starting character position of the bookmark.
@property (readonly) MicrosoftWordE160 storyType;  // Returns the story type for the bookmark.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's referred to by the bookmark.

- (MicrosoftWordBookmark *) copyBookmarkName:(NSString *)name NS_RETURNS_NOT_RETAINED;  // Sets the bookmark specified by the name argument to the location marked by another bookmark, and returns a bookmark object.

@end

// Represents options associated with border object.
@interface MicrosoftWordBorderOptions : MicrosoftWordBaseObject

@property BOOL alwaysInFront;  // Returns or sets if page borders are displayed in front of the document text.
@property MicrosoftWordE272 distanceFrom;  // Returns or sets a value that indicates whether the specified page border is measured from the edge of the page or from the text it surrounds.
@property NSInteger distanceFromBottom;  // Returns or sets the space in points between the text and the bottom border.
@property NSInteger distanceFromLeft;  // Returns or sets the space in points between the text and the left border.
@property NSInteger distanceFromRight;  // Returns or sets the space in points between the right edge of the text and the right border. 
@property NSInteger distanceFromTop;  // Returns or sets the space in points between the text and the top border.
@property BOOL enableBorders;  // Returns or sets border formatting for the specified object.
@property BOOL enableFirstPageInSection;  // Returns or sets if page borders are enabled for the first page in the section.
@property BOOL enableOtherPagesInSection;  // Returns or sets if page borders are enabled for all pages in the section except for the first page. 
@property (readonly) BOOL hasHorizontal;  // Returns true if a horizontal border can be applied to the object. 
@property (readonly) BOOL hasVertical;  // Returns true if a vertical border can be applied to the object. 
@property (copy) NSColor *insideColor;  // Returns or sets the RGB color of the inside borders
@property MicrosoftWordE110 insideColorIndex;  // Returns or sets the color index of the inside borders. 
@property MicrosoftWordMCoI insideColorThemeIndex;  // Returns or sets the color of the inside borders.
@property MicrosoftWordE167 insideLineStyle;  // Returns or sets the inside border for the specified object.
@property MicrosoftWordE168 insideLineWidth;  // Returns or sets the line width of the inside border of an object.
@property BOOL joinBorders;  // Returns or sets if vertical borders at the edges of paragraphs and tables are removed so that the horizontal borders can connect to the page border.
@property (copy) NSColor *outsideColor;  // Returns or sets the RGB color of the outside borders
@property MicrosoftWordE110 outsideColorIndex;  // Returns or sets the color index of the outside borders. 
@property MicrosoftWordMCoI outsideColorThemeIndex;  // Returns or sets the color of the outside borders.
@property MicrosoftWordE167 outsideLineStyle;  // Returns or sets the outside border for the specified object.
@property MicrosoftWordE168 outsideLineWidth;  // Returns or sets the line width of the outside border of an object.
@property BOOL shadow;  // Returns or sets if the specified border is formatted as shadowed.
@property BOOL surroundFooter;  // Returns or sets if a page border encompasses the document footer.
@property BOOL surroundHeader;  // Returns or sets if a page border encompasses the document header.

- (void) applyPageBordersToAllSections;  // Applies the specified page-border formatting to all sections in a document.

@end

// Represents a border of an object.
@interface MicrosoftWordBorder : MicrosoftWordBaseObject

@property MicrosoftWordE271 artStyle;  // Returns or sets the graphical page-border design for a document.
@property NSInteger artWidth;  // Returns or sets the width in points of the specified graphical page border.
@property (copy) NSColor *color;  // Returns or sets the RGB color for the specified border object. 
@property MicrosoftWordE110 colorIndex;  // Returns or sets the color index for the specified border.
@property MicrosoftWordMCoI colorThemeIndex;  // Returns or sets the color for the specified border or font object.
@property (readonly) BOOL inside;  // Returns true if an inside border can be applied to the specified object.
@property MicrosoftWordE167 lineStyle;  // Returns or sets the border line style for the specified object.
@property MicrosoftWordE168 lineWidth;  // Returns or sets the line width of an object's border.
@property BOOL visible;  // Returns or sets if the border object is visible


@end

// Represents the browser tool used to move the insertion point to objects in a document. This tool is comprised of the three buttons at the bottom of the vertical scroll bar.
@interface MicrosoftWordBrowser : MicrosoftWordBaseObject

@property MicrosoftWordE206 browserTarget;  // Returns or sets the document item that the previous and next methods locate.

- (void) nextForBrowser;  // Moves the selection to the next item indicated by the browser target. Use the browser target property to change the browser target.
- (void) previousForBrowser;  // Moves the selection to the previous item indicated by the browser target. Use the browser target property to change the browser target.

@end

@interface MicrosoftWordBuildingBlockCategory : MicrosoftWordBaseObject

- (SBElementArray *) buildingBlocks;

@property (copy, readonly) MicrosoftWordBuildingBlockType *buildingBlockType;  // Returns a building block type object that represents the type of building block for a building block category.
@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *name;  // Returns the name of the specified object.

- (MicrosoftWordBuildingBlock *) addBuildingBlockToCategoryName:(NSString *)name fromLocation:(MicrosoftWordTextRange *)fromLocation objectDescription:(NSString *)objectDescription insertOptions:(MicrosoftWordE330)insertOptions;  // Creates a new building block.

@end

// Represents a type of building block.
@interface MicrosoftWordBuildingBlockType : MicrosoftWordBaseObject

- (SBElementArray *) buildingBlockCategories;

@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *name;  // Returns a String that represents the localized name of a building block type.


@end

@interface MicrosoftWordBuildingBlock : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordBuildingBlockType *buildingBlockType;  // Returns a building block type object that represents the type for a building block.
@property (copy, readonly) MicrosoftWordBuildingBlockCategory *category;  // Returns a Category object that represents the category for a building block.
@property (copy) NSString *objectDescription;  // Gets or sets a string that represents the description for a building block.
@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *identifier;  // Returns a string that represents the internal identification number for a building block.
@property NSInteger insertOptions;  // Gets or sets an integer that represents how to insert the contents of a building block into a document.
@property (copy) NSString *name;  // Gets or sets a string that represents the name of a building block.
@property (copy) NSString *value;  // Gets or sets a string that represents the contents of a building block.

- (MicrosoftWordTextRange *) insertBuildingBlockInLocation:(MicrosoftWordTextRange *)inLocation asRichText:(BOOL)asRichText;  // Inserts the value of a building block into a document and returns a text range object that represents the contents of the building block within the document.

@end

// Represents a single caption label.
@interface MicrosoftWordCaptionLabel : MicrosoftWordBaseObject

@property (readonly) BOOL builtIn;  // Returns true if this is a built-in caption label.
@property (readonly) MicrosoftWordE210 captionLabelId;  // Returns a constant that represents the type for the specified caption label if the built in property of the caption label object is true.
@property MicrosoftWordE117 captionLabelPosition;  // Returns or sets the position of caption label text.
@property NSInteger chapterStyleLevel;  // Returns or sets the heading style that marks a new chapter when chapter numbers are included with the specified caption label. The number 1 corresponds to Heading 1, 2 corresponds to Heading 2, and so on.
@property BOOL includeChapterNumber;  // Returns or sets if a chapter number is included with page numbers or a caption label.
@property (copy, readonly) NSString *name;  // Returns the name for the caption label
@property MicrosoftWordE153 numberStyle;  // Returns or sets the number style for the caption label object.
@property MicrosoftWordE120 separator;  // Returns or sets the character between the chapter number and the sequence number.


@end

// Represents a single check box form field.
@interface MicrosoftWordCheckBox : MicrosoftWordBaseObject

@property BOOL autoSize;  // True sizes the check box according to the font size of the surrounding text. False sizes the check box according to the size property.
@property BOOL checkBoxDefault;  // Returns or sets the default check box value. True if the default value is checked. 
@property BOOL checkBoxValue;  // Returns or sets if the check box is selected.  True if the check box is selected.
@property double checkboxSize;  // Returns or sets the size of the specified check box in points.
@property (readonly) BOOL valid;  // Returns if the check box object is valid.


@end

// Represents a co-authoring lock
@interface MicrosoftWordCoauthLock : MicrosoftWordBaseObject

@property BOOL isHeader_footer_lock;  // returns if this lock is a header/footer lock.
@property (copy, readonly) MicrosoftWordCoauthor *lockOwner;  // Get the owner of the co-authoring lock.
@property MicrosoftWordE332 lockType;  // Get the type of the co-authoring lock.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Gets a text range object that represents the portion of a document that contains the co-authoring lock.

- (void) unlock;  // Remove the co-authoring lock.

@end

// Represents a co-authoring update
@interface MicrosoftWordCoauthUpdate : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Gets a text range object that represents the portion of a document that contains the co-authoring update.


@end

// Represents a coauthor
@interface MicrosoftWordCoauthor : MicrosoftWordBaseObject

- (SBElementArray *) coauthLocks;

@property (copy, readonly) NSString *coauthorEmailAddress;  // The email address of coauthor.
@property (copy, readonly) NSString *coauthorId;  // The ID of coauthor.
@property (copy, readonly) NSString *coauthorName;  // The name of coauthor.
@property BOOL isMe;  // returns if this coauthor is me.


@end

// Represents a coauthoring
@interface MicrosoftWordCoauthoring : MicrosoftWordBaseObject

- (SBElementArray *) coauthors;
- (SBElementArray *) coauthLocks;
- (SBElementArray *) coauthUpdates;
- (SBElementArray *) conflicts;

@property BOOL canMerge;  // returns if coauthoring can merge.
@property BOOL canShare;  // returns if coauthoring can share.
@property (copy) MicrosoftWordCoauthor *myself;  // returns me as author.
@property BOOL pendingUpdates;  // returns if any updates are pending.


@end

// Represents a conflict
@interface MicrosoftWordConflict : MicrosoftWordBaseObject

@property (readonly) MicrosoftWordE216 conflictType;  // Returns the revision type.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the conflict


@end

// Represents a custom mailing label.
@interface MicrosoftWordCustomLabel : MicrosoftWordBaseObject

@property (readonly) BOOL dotMatrix;  // True if the printer type for the specified custom label is dot matrix. False if the printer type is either laser or ink jet.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property double height;  // Returns or sets the height of the object.
@property double horizontalPitch;  // Returns or sets the horizontal distance in points between the left edge of one custom mailing label and the left edge of the next mailing label.
@property (copy) NSString *name;  // Returns or set the name of the custom mailing label.
@property NSInteger numberAcross;  // Returns or sets the number of custom mailing labels across a page.
@property NSInteger numberDown;  // Returns or sets the number of custom mailing labels down the length of a page.
@property MicrosoftWordE233 pageSize;  // Returns or sets the page size for the specified custom mailing label.
@property double sideMargin;  // Returns or sets the side margin widths in points for the specified custom mailing label.
@property double topMargin;  // Returns or sets the distance in points between the top edge of the page and the top boundary of the body text.
@property (readonly) BOOL valid;  // True if the various properties for example, height, width, and number down for the specified custom label work together to produce a valid mailing label.
@property double verticalPitch;  // Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label.
@property double width;  // Returns or sets the width of the object.


@end

// Represents a single data merge field in a data source.
@interface MicrosoftWordDataMergeDataField : MicrosoftWordBaseObject

@property (copy, readonly) NSString *dataMergeDataFieldValue;  // Returns the contents of the mail merge data field or mapped data field for the current record. Use the active record property to set the active record in a data merge data source.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the data merge data field object.


@end

// Represents the data merge data source in a data merge operation.
@interface MicrosoftWordDataMergeDataSource : MicrosoftWordBaseObject

- (SBElementArray *) dataMergeFieldNames;
- (SBElementArray *) dataMergeDataFields;

@property MicrosoftWordE193 activeRecord;  // Returns back the index of the current active record or an enumeration specifying the record
@property (copy, readonly) NSString *connectString;  // Returns the connection string for the specified data merge data source.
@property NSInteger firstRecord;  // Returns or sets the number of the first data record to be merged in a data merge operation. 
@property (copy, readonly) NSString *headerSourceName;  // Returns the path and file name of the header source attached to the specified mail merge main document.
@property (readonly) MicrosoftWordE195 headerSourceType;  // Returns a value that indicates the way the header source is being supplied for the mail merge operation.
@property NSInteger lastRecord;  // Returns or sets the number of the last data record to be merged in a data merge operation. 
@property (readonly) MicrosoftWordE195 mailMergeDataSourceType;  // Returns the type of data merge data source.
@property (copy, readonly) NSString *name;  // Returns the name of the data merge data source.
@property (copy) NSString *queryString;  // Returns or sets the query string, SQL statement, used to retrieve a subset of the data in a data merge data source.

- (BOOL) findRecordFindText:(NSString *)findText fieldName:(NSString *)fieldName;  // Searches the contents of the specified data merge data source for text in a particular field. Returns true if the search text is found.

@end

// Represents a data merge field name in a data source.
@interface MicrosoftWordDataMergeFieldName : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the data merge field name object


@end

// Represents a single data merge field in a document.
@interface MicrosoftWordDataMergeField : MicrosoftWordBaseObject

@property (copy) MicrosoftWordTextRange *dataMergeFieldRange;  // Returns or sets a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters.
@property (readonly) MicrosoftWordE183 formFieldType;  // The type of this data merge field
@property BOOL locked;  // Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results.
@property (copy, readonly) MicrosoftWordDataMergeField *nextDataMergeField;  // Returns the next data merge field
@property (copy, readonly) MicrosoftWordDataMergeField *previousMakeMergeField;  // Returns the previous data merge field


@end

// Represents the data merge functionality in Word.
@interface MicrosoftWordDataMerge : MicrosoftWordBaseObject

- (SBElementArray *) dataMergeFields;

@property (copy, readonly) MicrosoftWordDataMergeDataSource *dataSource;  // Returns or sets the destination of the data merge results.
@property MicrosoftWordE192 destination;  // Returns or sets the destination of the data merge results.
@property (copy) NSString *mailAddressFieldName;  // Returns or sets the name of the field that contains e-mail addresses that are used when the data merge destination is electronic mail.
@property BOOL mailAsAttachment;  // Returns or sets if the merge documents are sent as attachments when the data merge destination is an e-mail message or a fax.
@property (copy) NSString *mailSubject;  // Returns or sets the subject line used when the data merge destination is electronic mail.
@property MicrosoftWordE190 mainDocumentType;  // Returns or sets the data merge main document type.
@property (readonly) MicrosoftWordE191 state;  // Returns the current state of a data merge operation.
@property BOOL suppressBlankLines;  // Returns or sets if blank lines are suppressed when data merge fields in a data merge main document are empty.
@property BOOL viewDataMergeFieldCodes;  // Returns or sets if merge field names are displayed in a data merge main document. False if information from the current data record is displayed. 

- (void) check;  // Simulates the data merge operation, pausing to report each error as it occurs.
- (void) createDataSourceName:(NSString *)name passwordDocument:(NSString *)passwordDocument writePassword:(NSString *)writePassword headerRecord:(NSString *)headerRecord MSQuery:(BOOL)MSQuery SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1 connection:(NSString *)connection linkToSource:(BOOL)linkToSource;  // Creates a Word document that uses a table to store data for a data merge. The new data source is attached to the specified document, which becomes a main document if it's not one already.
- (void) createHeaderSourceName:(NSString *)name passwordDocument:(NSString *)passwordDocument writePassword:(NSString *)writePassword headerRecord:(NSString *)headerRecord;  // Creates a Word document that stores a header record that's used in place of the data source header record in a mail merge. This method attaches the new header source to the specified document, which becomes a main document if it's not one already.
- (void) editDataSource;  // Opens or switches to the mail merge data source.
- (void) editHeaderSource;  // Opens the header source attached to a mail merge main document, or activates the header source if it's already open.
- (void) editMainDocument;  // Activates the mail merge main document associated with the specified header source or data source document.
- (void) executeDataMergePause:(BOOL)pause;  // Performs the specified data merge operation. Performs the specified data merge operation.
- (MicrosoftWordDataMergeField *) makeNewDataMergeAskFieldTextRange:(MicrosoftWordTextRange *)textRange name:(NSString *)name prompt:(NSString *)prompt defaultAskText:(NSString *)defaultAskText askOnce:(BOOL)askOnce;  // Create a new data merge ask field
- (MicrosoftWordDataMergeField *) makeNewDataMergeFillInFieldTextRange:(MicrosoftWordTextRange *)textRange prompt:(NSString *)prompt defaultFillInText:(NSString *)defaultFillInText askOnce:(BOOL)askOnce;  // Create a new data merge fill in field
- (MicrosoftWordDataMergeField *) makeNewDataMergeIfFieldTextRange:(MicrosoftWordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(MicrosoftWordE196)comparison compareTo:(NSString *)compareTo trueText:(NSString *)trueText falseText:(NSString *)falseText;  // Create a new data merge if field
- (MicrosoftWordDataMergeField *) makeNewDataMergeNextFieldTextRange:(MicrosoftWordTextRange *)textRange;  // Create a new data merge next field
- (MicrosoftWordDataMergeField *) makeNewDataMergeNextIfFieldTextRange:(MicrosoftWordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(MicrosoftWordE196)comparison compareTo:(NSString *)compareTo;  // Create a new data merge next if field
- (MicrosoftWordDataMergeField *) makeNewDataMergeRecFieldTextRange:(MicrosoftWordTextRange *)textRange;  // Create a new data merge rec field
- (MicrosoftWordDataMergeField *) makeNewDataMergeSequenceFieldTextRange:(MicrosoftWordTextRange *)textRange;  // Create a new data merge sequence field
- (MicrosoftWordDataMergeField *) makeNewDataMergeSetFieldTextRange:(MicrosoftWordTextRange *)textRange name:(NSString *)name valueText:(NSString *)valueText;  // Create a new data merge set field
- (MicrosoftWordDataMergeField *) makeNewDataMergeSkipIfFieldTextRange:(MicrosoftWordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(MicrosoftWordE196)comparison compareTo:(NSString *)compareTo;  // Create a new data merge skip if field
- (void) openDataSourceName:(NSString *)name format:(MicrosoftWordE162)format confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly linkToSource:(BOOL)linkToSource addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate connection:(NSString *)connection SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1;  // Attaches a data source to the specified document, which becomes a main document if it's not one already.
- (void) openHeaderSourceName:(NSString *)name format:(MicrosoftWordE162)format confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate;  // Attaches a mail merge header source to the specified document.
- (void) useAddressBookBookType:(NSString *)bookType;  // Selects the address book that's used as the data source for a mail merge operation.

@end

// Contains global application-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page. You can return or set attributes either at the application global level or at the document level.
@interface MicrosoftWordDefaultWebOptions : MicrosoftWordBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page.
@property BOOL alwaysSaveInDefaultEncoding;  // Returns or saves if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.  The default value is False.
@property BOOL checkIfOfficeIsHtmleditor;  // Returns or sets if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word.
@property BOOL checkIfWordIsDefaultHtmleditor;  // Returns or sets if Microsoft Word checks to see whether it is the default HTML editor when you start Word. The default value is true.
@property MicrosoftWordMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property MicrosoftWordMSsz screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL updateLinksOnSave;  // Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved.
@property BOOL useLongFileNames;  // Returns or sets if long file names are used when you save the document as a Web page.


@end

// Represents a built-in dialog box.
@interface MicrosoftWordDialog : MicrosoftWordBaseObject

@property MicrosoftWordE185 defaultDialogTab;  // Returns or sets the active tab when the specified dialog box is displayed.
@property (readonly) MicrosoftWordE186 dialogType;  // The built-in dialog this object represents.

- (NSInteger) displayWordDialogTimeOut:(NSInteger)timeOut;  // Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a Long that indicates which button was clicked to close the dialog box.
- (void) executeDialog;  // Applies the current settings of a Microsoft Word dialog box.
- (NSInteger) showTimeOut:(NSInteger)timeOut;  // Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a number which indicates the button used to dismiss the dialog box.

@end

// Represents a single version of a document.
@interface MicrosoftWordDocumentVersion : MicrosoftWordBaseObject

@property (copy, readonly) NSString *comment;  // Returns the comment associated with the specified version of a document.
@property (copy, readonly) NSDate *dateValue;  // The date and time that the document version was saved.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *savedBy;  // Returns the name of the user who saved the specified version of the document.

- (void) openVersion;  // Opens the specified version of the document.

@end

// Represents a Microsoft Word document.
@interface MicrosoftWordDocument : MicrosoftWordBaseObject

- (SBElementArray *) documentProperties;
- (SBElementArray *) customDocumentProperties;
- (SBElementArray *) bookmarks;
- (SBElementArray *) tables;
- (SBElementArray *) footnotes;
- (SBElementArray *) endnotes;
- (SBElementArray *) WordComments;
- (SBElementArray *) sections;
- (SBElementArray *) paragraphs;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) characters;
- (SBElementArray *) fields;
- (SBElementArray *) formFields;
- (SBElementArray *) WordStyles;
- (SBElementArray *) frames;
- (SBElementArray *) tablesOfFigures;
- (SBElementArray *) variables;
- (SBElementArray *) revisions;
- (SBElementArray *) tablesOfContents;
- (SBElementArray *) tablesOfAuthorities;
- (SBElementArray *) windows;
- (SBElementArray *) indexes;
- (SBElementArray *) subdocuments;
- (SBElementArray *) storyRanges;
- (SBElementArray *) hyperlinkObjects;
- (SBElementArray *) shapes;
- (SBElementArray *) listTemplates;
- (SBElementArray *) WordLists;
- (SBElementArray *) inlineShapes;
- (SBElementArray *) documentVersions;
- (SBElementArray *) readabilityStatistics;
- (SBElementArray *) grammaticalErrors;
- (SBElementArray *) spellingErrors;
- (SBElementArray *) mathObjects;

@property (copy, readonly) MicrosoftWordWindow *activeWindow;  // Returns a window object that represents the active window, the window with the focus. If there are no windows open, an error occurs.
@property (copy) MicrosoftWordTemplate *attachedTemplate;  // Returns of sets a template object that represents the template attached to the specified document. To set this property, specify either the name of the template or a template object.
@property BOOL autoHyphenation;  // Returns or sets if automatic hyphenation is turned on for the specified document.
@property (copy) MicrosoftWordShape *backgroundShape;  // Returns or sets a shape object that represents the background image for the specified document.
@property MicrosoftWordE184 clickAndTypeParagraphStyle;  // Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document.
@property (copy, readonly) MicrosoftWordCoauthoring *coauthoring;
@property MicrosoftWordE314 compatibleVersion;  // Returns or sets the compatibility options to a given application version.
@property NSInteger consecutiveHyphensLimit;  // Returns or sets the maximum number of consecutive lines that can end with hyphens. If this property is set to zero, any number of consecutive lines can end with hyphens.
@property (copy, readonly) MicrosoftWordDataMerge *dataMerge;  // Returns the data merge object.
@property double defaultTabStop;  // Returns or sets the interval in points between the default tab stops in the specified document.
@property (readonly) MicrosoftWordE184 defaultTableStyle;  // Returns the default table style for this document.
@property (copy, readonly) MicrosoftWordOfficeTheme *documentTheme;  // Returns an office theme object that represents the Microsoft Office theme applied to a document.
@property (readonly) MicrosoftWordE222 document_type;  // Returns the document type.
@property MicrosoftWordE108 eastAsianLineBreak;  // Returns or sets the line break control level for the specified document. This property is ignored if the “east asian line break control” property is set to false. Note that “east asian line break control” is a paragraph format property.
@property BOOL embedTrueTypeFonts;  // Returns or set if Microsoft Word embeds TrueType fonts in a document when it's saved. This allow others to view the document with the same fonts that were used to create it.
@property (copy, readonly) MicrosoftWordEndnoteOptions *endnoteOptions;  // Returns the endnote options object.
@property (copy, readonly) MicrosoftWordEnvelope *envelopeObject;  // Returns the envelop object.
@property (copy, readonly) MicrosoftWordFieldOptions *fieldOptions;
@property (copy, readonly) MicrosoftWordFootnoteOptions *footnoteOptions;  // Returns the footnote options object.
@property (copy, readonly) NSString *fullName;  // Returns the full name of the document.
@property BOOL grammarChecked;  // True if a grammar check has been run on the document. False if some of the document hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false.
@property double gridDistanceHorizontal;  // Returns or sets the amount of horizontal space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize shape object or east asian characters in the document.
@property double gridDistanceVertical;  // Returns or sets the amount of vertical space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize shape object or east asian characters in the document.
@property BOOL gridOriginFromMargin;  // Returns or sets if Microsoft Word starts the character grid from the upper-left corner of the page.
@property double gridOriginHorizontal;  // Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing shape object or east asian characters to begin in the document.
@property double gridOriginVertical;  // Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing shape object or east asian characters to begin in the document.
@property NSInteger gridSpaceBetweenHorizontalLines;  // Returns or sets the interval at which Microsoft Word displays horizontal character gridlines in print layout view.
@property NSInteger gridSpaceBetweenVerticalLines;  // Returns or sets the interval at which Microsoft Word displays vertical character gridlines in print layout view.
@property (readonly) BOOL hasPassword;  // True if a password is required to open the specified document.
@property BOOL hyphenateCaps;  // Returns or sets if words in all capital letters can be hyphenated.
@property NSInteger hyphenationZone;  // Returns or sets the width of the hyphenation zone, in points. The hyphenation zone is the maximum amount of space that Microsoft Word leaves between the end of the last word in a line and the right margin.
@property (readonly) BOOL integralsUsesSubscripts;  // Gets or sets a value that specifies the default location of limits for integrals.
@property (readonly) BOOL isMasterDocument;  // True if the specified document is a master document. A master document includes one or more subdocuments.
@property (readonly) BOOL isSubdocument;  // True if the specified document is opened in a separate document window as a subdocument of a master document.
@property MicrosoftWordE107 justificationMode;  // Returns or sets the character spacing adjustment for the specified document.
@property (copy) MicrosoftWordLetterContent *letterContent;  // Return or sets the letter content object associated with the document.
@property (readonly) MicrosoftWordE327 mathBinaryOperatorBreak;  // Gets or sets a value that specifies where Microsoft Word places binary operators when equations span two or more lines.
@property (readonly) MicrosoftWordE326 mathDefaultJustification;  // Gets or sets a value that indicates the default justification.
@property (copy, readonly) NSString *mathFontName;  // Gets or sets the name of the font that is used in a document to display equations.
@property (readonly) double mathLeftMargin;  // Gets or sets a value that specifies the left margin for equations.
@property (readonly) double mathRightMargin;  // Gets or sets a value that specifies the right margin for equations.
@property (readonly) MicrosoftWordE328 mathSubtractionOperator;  // Gets or sets a value that specifies how Microsoft Word handles a subtraction operator that falls before a line break.
@property (readonly) double mathWrap;  // Gets or sets a value that specifies the placement of the second line of an equation that wraps to a new line.
@property (copy, readonly) NSString *name;  // Returns the name of the document.
@property (readonly) BOOL naryUsesSubscripts;  // Gets or sets a value that specifies the default location of limits for n-ary objects other than integrals
@property (copy) NSString *noLineBreakAfter;  // Returns or sets the kinsoku characters after which Microsoft Word will not break a line.
@property (copy) NSString *noLineBreakBefore;  // Returns or sets the kinsoku characters before which Microsoft Word will not break a line.
@property (copy) MicrosoftWordPageSetup *pageSetup;  // Returns or sets the page setup object.
@property (copy) NSString *password;  // Sets a password that must be supplied to open the specified document. This is write-only property
@property (copy, readonly) NSString *path;  // Returns the path to the document.
@property BOOL printFormsData;  // Returns or sets if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form.
@property BOOL printPostScriptOverText;  // Returns or sets if PRINT field instructions such as PostScript commands in a document are to be printed on top of text and graphics when a PostScript printer is used.
@property BOOL printRevisions;  // Returns or sets if revision marks are printed with the document. False if revision marks aren't printed that is, tracked changes are printed as if they'd been accepted.
@property (readonly) MicrosoftWordE234 protectionType;  // Returns the protection type for the specified document.
@property (readonly) BOOL readOnly;  // True if changes to the document cannot be saved to the original document.
@property BOOL readOnlyRecommended;  // Returns or set if Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only.
@property BOOL removeDateAndTime;  // Returns or sets if Microsoft Word removes date and time from revisions upon saving a document.
@property BOOL removePersonalInformation;  // Returns or sets if Microsoft Word removes all user information from comments, revisions, and the properties dialog box upon saving a document.
@property (readonly) MicrosoftWordE161 saveFormat;  // Returns the file format of the specified document or file converter. Will be a unique number that specifies an external file converter or a constant.
@property BOOL saveFormsData;  // Returns or sets if Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database.
@property BOOL saveSubsetFonts;  // Returns or sets if Microsoft Word saves a subset of the embedded TrueType fonts with the document.
@property BOOL saveVersionsOnClose;  // Sets or returns whether or not versions are automatically saved when a document is closed.
@property BOOL saved;  // Returns or set the saved state. True if the specified document or template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed.
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showGrammaticalErrors;  // Returns or sets if grammatical errors are marked by a wavy green line in the document. To view grammatical errors in your document, you must set the check grammar as you type property of the Word options class to true.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are shown.
@property BOOL showRevisions;  // Returns or sets if tracked changes in the specified document are shown on the screen.
@property BOOL showSpellingErrors;  // Returns or sets if Microsoft Word underlines spelling errors in the document.  To view spelling errors in your document, you must set the check grammar as you type property of the Word options class to true.
@property BOOL snapToGrid;  // Returns or sets if shape object or east asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in the specified document.
@property BOOL snapToShapes;  // Returns or sets if Microsoft Word automatically aligns shape object or east asian characters with invisible gridlines that go through the vertical and horizontal edges of other shape object or east asian characters in the document.
@property BOOL spellingChecked;  // True if a spelling check has been run on the document. False if some of the document hasn't been checked for spelling. To see if the document contains spelling errors, get the count of spelling errors for the document.
@property BOOL subdocumentsExpanded;  // Returns or set if the subdocuments in the document are expanded.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a range object that represents the main document story.
@property BOOL trackRevisions;  // Returns or sets if changes are tracked in the document.
@property (copy, readonly) NSArray *unavailableFonts;  // Returns a list of fonts used in the document that are not available on the current system.
@property BOOL updateStylesOnOpen;  // Returns or sets if the styles in the specified document are updated to match the styles in the attached template each time the document is opened.
@property (readonly) BOOL useDefaultMathSettings;  // Gets or sets a value that indicates whether to use the default math settings when creating new equations.
@property (readonly) BOOL useSmallFractions;  // Gets or sets a value that indicates whether to use small fractions in equations contained within the document.
@property (copy, readonly) MicrosoftWordWebOptions *webOptions;  // Returns the web options object.
@property (copy) NSString *writePassword;  // Sets a password for saving changes to the specified document. This is write-only property
@property (readonly) BOOL writeReserved;  // True if the specified document is protected with a write password.

- (void) acceptAllRevisions;  // Accepts all tracked changes in the document.
- (void) acceptAllShownRevisions;  // Accepts all shown tracked changes.
- (void) applyDocumentThemeFileName:(NSString *)fileName;  // Applies an Office theme to the document.
- (void) autoFormat;  // Automatically formats a document
- (BOOL) canCheckIn;  // Returns True if Word can check in a specified document to a server.
- (void) changeDefaultTableStyleStyle:(MicrosoftWordE184)style changeInTemplate:(BOOL)changeInTemplate;  // Sets the default table style.
- (void) checkConsistency;  // Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a document from a local computer to a server, and sets the local document to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(MicrosoftWordE331)versionType;  // Returns a document from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) closePrintPreview;  // Switches the specified document from print preview to the previous view. If the specified document isn't in print preview, an error occurs.
- (void) comparePath:(NSString *)path authorName:(NSString *)authorName target:(MicrosoftWordE297)target detectFormatChanges:(BOOL)detectFormatChanges ignoreAllComparisonWarnings:(BOOL)ignoreAllComparisonWarnings addToRecentFiles:(BOOL)addToRecentFiles;  // Compares this document with another.
- (NSInteger) computeStatisticsStatistic:(MicrosoftWordE155)statistic includeFootnotesAndEndnotes:(BOOL)includeFootnotesAndEndnotes;  // Computes a set of readability statistics for this document.  This must be called before accessing the readability statistics for this document.
- (void) copyStylesFromTemplateTemplate:(NSString *)template_;// NS_RETURNS_NOT_RETAINED;  // Copies styles from the specified template to a document.
- (MicrosoftWordLetterContent *) createLetterContentDateFormat:(NSString *)dateFormat includeHeaderFooter:(BOOL)includeHeaderFooter pageDesign:(NSString *)pageDesign letterStyle:(MicrosoftWordE245)letterStyle letterhead:(BOOL)letterhead letterheadLocation:(MicrosoftWordE246)letterheadLocation letterheadSize:(double)letterheadSize recipientName:(NSString *)recipientName recipientAddress:(NSString *)recipientAddress salutation:(NSString *)salutation salutationType:(MicrosoftWordE247)salutationType recipientReference:(NSString *)recipientReference mailingInstructions:(NSString *)mailingInstructions attentionLine:(NSString *)attentionLine subject:(NSString *)subject ccList:(NSString *)ccList returnAddress:(NSString *)returnAddress senderName:(NSString *)senderName closing:(NSString *)closing senderCompany:(NSString *)senderCompany senderJobTitle:(NSString *)senderJobTitle senderInitials:(NSString *)senderInitials enclosureCount:(NSInteger)enclosureCount;  // Create a new letter content object for use with the letter wizard.
- (MicrosoftWordTextRange *) createRangeStart:(NSInteger)start end:(NSInteger)end;  // Returns a text range object by using the specified starting and ending character positions.
- (void) dataForm;  // Displays the data form dialog box, in which you can modify data records. You can use this method with a mail merge main document, a mail merge data source, or any document that contains data delimited by table cells or separator characters
- (void) deleteAllComments;  // Deletes all comments.
- (void) deleteAllShownComments;  // Deletes all shown comments.
- (void) fitToPages;  // Decreases the font size of text just enough so that the document will fit on one fewer pages. An error occurs if Word is unable to reduce the page count by one.
- (void) followHyperlinkAddress:(NSString *)address subAddress:(NSString *)subAddress newWindow:(BOOL)newWindow addHistory:(BOOL)addHistory extraInfo:(NSString *)extraInfo;  // This method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application. If the hyperlink uses the file protocol, this method opens the document instead of downloading it.
- (NSString *) getActiveWritingStyleLanguageID:(MicrosoftWordE182)languageID;  // Returns the writing style for a specified language.
- (NSArray *) getCrossReferenceItemsReferenceType:(MicrosoftWordE211)referenceType;  // Returns an list of items that can be cross-referenced based on the specified cross-reference type.
- (BOOL) getDocumentCompatibilityCompatibilityItem:(MicrosoftWordE231)compatibilityItem;  // Returns the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.
- (MicrosoftWordStoryRange *) getStoryRangeStoryType:(MicrosoftWordE160)storyType;  // Returns a range object that represents the story specified by the story type argument.
- (void) makeCompatibilityDefault;  // Uses the correct settings of the document compatibility options set by the set compatibility options method as the default for new documents.
- (void) manualHyphenation;  // Initiates manual hyphenation of a document, one line at a time. The user is prompted to accept or decline suggested hyphenations.
- (MicrosoftWordField *) markEntryForTableOfContentsRange:(MicrosoftWordTextRange *)range entry:(NSString *)entry tableID:(NSString *)tableID level:(NSInteger)level;  // Inserts a table of contents entry field after the specified range. The method returns a field object representing the new field.
- (MicrosoftWordField *) markEntryForTableOfFiguresRange:(MicrosoftWordTextRange *)range entry:(NSString *)entry tableID:(NSString *)tableID level:(NSInteger)level;  // Inserts a table of figures entry field after the specified range. The method returns a field object representing the new field.
- (MicrosoftWordField *) markForIndexRange:(MicrosoftWordTextRange *)range entry:(NSString *)entry crossReference:(NSString *)crossReference bookmarkName:(NSString *)bookmarkName;  // Inserts an index entry field after the specified range. The method returns a field object representing the new field.
- (void) mergeFileName:(NSString *)fileName;  // Merges the changes marked with revision marks from one document to another.
- (void) mergeSubdocumentsFirstSubdocument:(MicrosoftWordSubdocument *)firstSubdocument lastSubdocument:(MicrosoftWordSubdocument *)lastSubdocument;  // Merges the specified subdocuments of a master document into a single subdocument.
- (void) presentIt;  // Opens PowerPoint with the specified Word document loaded.
- (void) printPreview;  // Switches the view to print preview.
- (void) protectProtectionType:(MicrosoftWordE234)protectionType noReset:(BOOL)noReset password:(NSString *)password useIrm:(BOOL)useIrm enforceStyleLocks:(BOOL)enforceStyleLocks;  // Helps to protect the specified document from changes. When a document is protected, users can make only limited changes, such as adding annotations, making revisions, or completing a form.
- (BOOL) redoTimes:(NSInteger)times;  // Redoes the last action that was undone. It reverses the undo method. Returns true if the actions were redone successfully.
- (void) rejectAllRevisions;  // Rejects all tracked changes in the document.
- (void) rejectAllShownRevisions;  // Rejects all shown tracked changes.
- (void) reload;  // Reloads a cached document by resolving the hyperlink to the document and downloading it.
- (void) removeReminder;  // Remove the document's reminder.
- (void) repaginate;  // Repaginates the entire document
- (void) runLetterWizardLetterContent:(MicrosoftWordLetterContent *)letterContent wizardMode:(BOOL)wizardMode;  // Runs the Letter Wizard on the document
- (void) saveAsFileName:(NSString *)fileName fileFormat:(MicrosoftWordE161)fileFormat lockComments:(BOOL)lockComments password:(NSString *)password addToRecentFiles:(BOOL)addToRecentFiles writePassword:(NSString *)writePassword readOnlyRecommended:(BOOL)readOnlyRecommended embedTruetypeFonts:(BOOL)embedTruetypeFonts saveNativePictureFormat:(BOOL)saveNativePictureFormat saveFormsData:(BOOL)saveFormsData textEncoding:(NSInteger)textEncoding insertLineBreaks:(BOOL)insertLineBreaks allowSubstitutions:(BOOL)allowSubstitutions lineEndingType:(MicrosoftWordE311)lineEndingType HTMLDisplayOnlyOutput:(BOOL)HTMLDisplayOnlyOutput maintainCompatibility:(BOOL)maintainCompatibility;  // Saves the document with a new name or format.
- (void) saveVersionComment:(NSString *)comment;  // Saves a version of the document with a comment.
- (void) sendHtmlMail;  // Opens a message window for sending the specified document, formatted as html, through Microsoft Entourage.
- (void) sendMail;  // Opens a message window for sending the specified document through your registered mail program.
- (void) setActiveWritingStyleLanguageID:(MicrosoftWordE182)languageID writingStyle:(NSString *)writingStyle;  // Sets the writing style for the specified language.
- (void) setDocumentCompatibilityCompatibilityItem:(MicrosoftWordE231)compatibilityItem isCompatible:(BOOL)isCompatible;  // Sets the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.
- (void) setReminderAt:(NSDate *)at;  // Create or modify a reminder for the document.
- (BOOL) undoTimes:(NSInteger)times;  // Undoes the last action or a sequence of actions, which are displayed in the undo list. Returns true if the actions were successfully undone.
- (void) undoClear;  // Clear the list of actions that can be undone.
- (void) unprotectPassword:(NSString *)password;  // Removes protection from the specified document. If the document isn't protected, this method generates an error.
- (void) updateStyles;  // Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.
- (void) upgrade;  // upgrade document
- (void) viewPropertyBrowser;  // Displays the property window for the selected control in the specified document.
- (void) webPagePreview;  // Displays a preview of the document as it would look if saved as a Web page.

@end

// Represents a dropped capital letter at the beginning of a paragraph.
@interface MicrosoftWordDropCap : MicrosoftWordBaseObject

@property double distanceFromText;  // Returns or sets the distance in points between the dropped capital letter and the paragraph text. 
@property MicrosoftWordE172 dropPosition;  // Returns or sets the position of a dropped capital letter.
@property (copy) NSString *fontName;  // Returns or sets the name of the font for the dropped capital letter.
@property NSInteger linesToDrop;  // Returns or sets the height in lines of the specified dropped capital letter.

- (void) enable;  // Formats the first character in the specified paragraph as a dropped capital letter.

@end

// Represents a drop-down form field that contains a list of items in a form.
@interface MicrosoftWordDropDown : MicrosoftWordBaseObject

- (SBElementArray *) listEntries;

@property NSInteger dropDownDefault;  // Returns or sets the default drop-down item. The first item in a drop-down form field is 1, the second item is 2, and so on.
@property NSInteger dropDownValue;  // Returns or sets the number of the selected item in a drop-down form field.
@property (readonly) BOOL valid;  // Returns if the drop down object is valid.


@end

// A representation of the options associated with endnotes.
@interface MicrosoftWordEndnoteOptions : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordTextRange *endnoteContinuationNotice;  // Returns a text range object that represents the endnote continuation notice.
@property (copy, readonly) MicrosoftWordTextRange *endnoteContinuationSeparator;  // Returns a text range object that represents the endnote continuation separator.
@property MicrosoftWordE175 endnoteLocation;  // Returns or sets the position of all endnotes.
@property MicrosoftWordE152 endnoteNumberStyle;  // Returns or sets the number style for endnotes.
@property MicrosoftWordE173 endnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks.
@property (copy, readonly) MicrosoftWordTextRange *endnoteSeparator;  // Returns a text range object that represents the endnote separator.
@property NSInteger endnoteStartingNumber;  // Returns or sets the starting endnote number.

- (void) endnoteConvert;  // Converts endnotes to footnotes, or vice versa.
- (void) swapWithFootnotes;  // Converts all footnotes in a document to endnotes and vice versa.

@end

// Represents an endnote.
@interface MicrosoftWordEndnote : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) MicrosoftWordTextRange *noteReference;  // Returns a text range object that represents a endnote mark.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the endnote object.


@end

// Represents an envelope.
@interface MicrosoftWordEnvelope : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordTextRange *address;  // Returns the envelope delivery address as a text range object.
@property double addressFromLeft;  // Returns or sets the distance in points between the left edge of the envelope and the delivery address.
@property double addressFromTop;  // Returns or sets the distance in points between the top edge of the envelope and the delivery address.
@property (copy, readonly) MicrosoftWordWordStyle *addressStyle;  // Returns a Word style object that represents the delivery address style for the envelope.
@property BOOL defaultFaceUp;  // Returns or sets if envelopes are fed face up by default.
@property double defaultHeight;  // Returns or sets the default envelope height, in points.
@property BOOL defaultOmitReturnAddress;  // Returns or sets if the return address is omitted from envelopes by default.
@property MicrosoftWordE244 defaultOrientation;  // Returns or sets the default orientation for feeding envelopes.
@property BOOL defaultPrintFIMA;  // Returns or sets if a Facing Identification Mark FIM-A to envelopes by default. A FIM-A code is used to presort courtesy reply mail. The default print barcode property must be set to true before this property is set.
@property BOOL defaultPrintBarCode;  // Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set
@property (copy) NSString *defaultSize;  // Returns or sets the default envelope size. If you set either the default height or default width property, the property  is automatically changed to return custom size.
@property double defaultWidth;  // Returns or sets the default envelope width, in points.
@property MicrosoftWordE207 feedSource;  // Returns or sets the paper tray for the envelope.
@property (copy, readonly) MicrosoftWordTextRange *returnAddress;  // Returns the envelope return address as a text range object.
@property double returnAddressFromLeft;  // Returns or sets the distance in points between the left edge of the envelope and the return address.
@property double returnAddressFromTop;  // Returns or sets the distance in points between the top edge of the envelope and the return address.
@property (copy, readonly) MicrosoftWordWordStyle *returnAddressStyle;  // Returns a Word style object that represents the return address style for the envelope.

- (void) insertEnvelopeDataExtractAddress:(BOOL)extractAddress address:(NSString *)address autoText:(NSString *)autoText omitReturnAddress:(BOOL)omitReturnAddress returnAddress:(NSString *)returnAddress returnAutotext:(NSString *)returnAutotext printBarCode:(BOOL)printBarCode printFIMA:(BOOL)printFIMA envelopeSize:(NSString *)envelopeSize envelopeHeight:(NSInteger)envelopeHeight envlopeWidth:(NSInteger)envlopeWidth feedSource:(BOOL)feedSource addressFromLeft:(NSInteger)addressFromLeft addressFromTop:(NSInteger)addressFromTop returnAddressFromLeft:(NSInteger)returnAddressFromLeft returnAddressFromTop:(NSInteger)returnAddressFromTop defaultFaceUp:(BOOL)defaultFaceUp defaultOrientation:(MicrosoftWordE244)defaultOrientation sizeFromPageSetup:(BOOL)sizeFromPageSetup showPageSetupDialog:(BOOL)showPageSetupDialog createNewDocument:(BOOL)createNewDocument;  // Inserts an envelope as a separate section at the beginning of the specified document.
- (void) printOutEnvelopeExtractAddress:(BOOL)extractAddress address:(NSString *)address autoText:(NSString *)autoText omitReturnAddress:(BOOL)omitReturnAddress returnAddress:(NSString *)returnAddress returnAutotext:(NSString *)returnAutotext printBarCode:(BOOL)printBarCode printFIMA:(BOOL)printFIMA envelopeSize:(NSString *)envelopeSize envelopeHeight:(NSInteger)envelopeHeight envlopeWidth:(NSInteger)envlopeWidth feedSource:(BOOL)feedSource addressFromLeft:(NSInteger)addressFromLeft addressFromTop:(NSInteger)addressFromTop returnAddressFromLeft:(NSInteger)returnAddressFromLeft returnAddressFromTop:(NSInteger)returnAddressFromTop defaultFaceUp:(BOOL)defaultFaceUp defaultOrientation:(MicrosoftWordE244)defaultOrientation sizeFromPageSetup:(BOOL)sizeFromPageSetup showPageSetupDialog:(BOOL)showPageSetupDialog;
- (void) updateDocument;  // Updates the envelope in the document with the current envelope settings.

@end

@interface MicrosoftWordFieldOptions : MicrosoftWordBaseObject

@property BOOL locked;


@end

// Represents a field.
@interface MicrosoftWordField : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) MicrosoftWordTextRange *fieldCode;  // Returns a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters.
@property (readonly) MicrosoftWordE187 fieldKind;  // Returns the type of link for a field object.
@property (copy) NSString *fieldText;  // Returns or sets data in an ADDIN field. The data is not visible in the field code or result. It is only accessible by returning the value of the data property. If the field isn't an ADDIN field, this property will cause an error.
@property (readonly) MicrosoftWordE183 fieldType;  // Returns the field type.
@property (copy, readonly) MicrosoftWordInlineShape *inlineShape;  // Returns the inline shape object associated with this field
@property (copy, readonly) MicrosoftWordLinkFormat *linkFormat;  // Returns the link format object associated with this field object.
@property BOOL locked;  // Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results.
@property (copy, readonly) MicrosoftWordField *nextField;  // Returns the next field object.
@property (copy, readonly) MicrosoftWordField *previousField;  // Returns the previous field object.
@property (copy) MicrosoftWordTextRange *resultRange;  // Returns or sets a text range object that represents a field's result. You can access a field result without changing the view from field codes.
@property BOOL showCodes;  // Returns or sets if field codes are displayed for the specified field instead of field results.

- (void) clickObject;  // Clicks the specified field. If the field is a GOTOBUTTON field, this method moves the insertion point to the specified location or selects the specified bookmark. If the field is a HYPERLINK field, this method jumps to the target location.
- (BOOL) updateField;  // Updates the result of the field object. When applied to a field object, returns true if the field is updated successfully.

@end

// Represents a file converter that's used to open or save files.
@interface MicrosoftWordFileConverter : MicrosoftWordBaseObject

@property (readonly) BOOL canOpen;  // Returns true if the specified file converter is designed to open files.
@property (readonly) BOOL canSave;  // Returns true if the specified file converter is designed to save files.
@property (copy, readonly) NSString *className;  // Returns a unique name that identifies the file converter.
@property (copy, readonly) NSString *extensions;  // Returns the file name extensions associated with the specified file converter object.
@property (copy, readonly) NSString *formatName;  // Returns the name of the specified file converter.
@property (copy, readonly) NSString *name;  // Returns the name of the file converter.
@property (readonly) NSInteger openFormat;  // Returns the file format of the specified file converter. It will be a unique number that represents an external file converter.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified file converter. 
@property (readonly) NSInteger saveFormat;  // Returns the file format of the specified document or file converter. It will be a unique number that specifies an external file converter.


@end

// Represents the criteria for a find operation.
@interface MicrosoftWordFind : MicrosoftWordBaseObject

@property (copy) NSString *content;  // Returns or sets the text in the find object.
@property (copy, readonly) MicrosoftWordFont *fontObject;  // Returns the font object associated with this find object.
@property BOOL format;  // Returns or set if formatting is included in the find operation.
@property BOOL forward;  // Returns or sets if the find operation searches forward through the document. False if it searches backward through the document.
@property (readonly) BOOL found;  // True if the search produces a match.
@property (copy, readonly) MicrosoftWordFrame *frame;  // Returns the frame object associated with the find object.
@property NSInteger highlight;  // Returns or sets if highlight formatting is included in the find criteria
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the find object
@property MicrosoftWordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property BOOL matchAllWordForms;  // Returns or sets if all forms of the text to find are found by the find operation for instance, if the text to find is sit, sat and sitting are found as well.
@property BOOL matchByte;  // Returns or sets if Microsoft Word distinguishes between full-width and half-width letters or characters during a search.
@property BOOL matchCase;  // Returns or sets if the find operation is case sensitive.
@property BOOL matchFuzzy;  // Returns or sets if Microsoft Word uses the nonspecific search options for Japanese text during a search.
@property BOOL matchSoundsLike;  // Returns or sets if words that sound similar to the text to find are returned by the find operation.
@property BOOL matchWholeWord;  // Returns or sets if the find operation locates only entire words and not text that's part of a larger word.
@property BOOL matchWildcards;  // Returns or sets if the text to find contains wildcards.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with the find object.
@property (copy, readonly) MicrosoftWordReplacement *replacement;  // Returns the replacement object associated with the find object.
@property MicrosoftWordE184 style;  // Returns or sets the Word style associated with the find object.
@property MicrosoftWordE182 supplementalLanguageID;  // Returns or sets the language for the text range object
@property MicrosoftWordE265 wrap;  // Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached or vice versa if forward is set to false or if the search text isn't found in the specified selection or range. 

- (void) clearAllFuzzyOptions;  // Clears all nonspecific search options associated with Japanese text.
- (MicrosoftWordEFRt) executeFindFindText:(NSString *)findText matchCase:(BOOL)matchCase matchWholeWord:(BOOL)matchWholeWord matchWildcards:(BOOL)matchWildcards matchSoundsLike:(BOOL)matchSoundsLike matchAllWordForms:(BOOL)matchAllWordForms matchForward:(BOOL)matchForward wrapFind:(MicrosoftWordE265)wrapFind findFormat:(BOOL)findFormat replaceWith:(NSString *)replaceWith replace:(MicrosoftWordE273)replace;  // Runs the specified find operation. Returns true if the find operation is successful.
- (void) setAllFuzzyOptions;  // Activates all nonspecific search options associated with Japanese text.

@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface MicrosoftWordFont : MicrosoftWordBaseObject

@property BOOL allCaps;  // Returns or sets if the font is formatted as all capital letters.
@property MicrosoftWordE124 animation;  // Returns or sets the type of animation applied to the font.
@property (copy) NSString *asciiName;  // Returns or sets the font used for Latin text characters with character codes from 0 through 127. 
@property BOOL bold;  // Returns or sets if the font is formatted as bold.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property (copy) NSColor *color;  // Returns or sets the RGB color of the font object.
@property MicrosoftWordE110 colorIndex;  // Returns or sets the color for the font object using an index.
@property MicrosoftWordMCoI colorThemeIndex;  // Returns or sets the color for the specified border or font object.
@property BOOL disableCharacterSpaceGrid;  // Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding font object.
@property BOOL doubleStrikeThrough;  // Returns or sets if the specified font is formatted as double strikethrough text.
@property (copy) NSString *eastAsianName;  // Returns or sets an East Asian font name.
@property BOOL emboss;  // Returns or sets if the specified font is formatted as embossed.
@property MicrosoftWordE114 emphasisMark;  // Returns or sets the emphasis mark for a character or designated character string.
@property BOOL engrave;  // Returns or sets if the specified font is formatted as engraved.
@property NSInteger fontPosition;  // Returns or sets the position of text in points relative to the base line. A positive number raises the text, and a negative number lowers it.
@property double fontSize;  // Returns or sets the font size.
@property BOOL hidden;  // Returns or sets if the font is formatted as hidden text.
@property BOOL italic;  // Returns or sets if the font is formatted as italic.
@property double kerning;  // Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically.
@property (copy) NSString *name;  // Returns or sets the font name associated with this font object.
@property (copy) NSString *otherName;  // Returns or sets the font used for characters with character codes from 128 through 255.
@property BOOL outline;  // Returns or sets if the specified font is formatted as outline.
@property NSInteger scaling;  // Returns or sets the scaling percentage applied to the font. This property stretches or compresses text horizontally as a percentage of the current size. The scaling range is from 1 through 600.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the font object.
@property BOOL shadow;  // Returns or sets if the specified font is formatted as shadowed.
@property BOOL smallCaps;  // Returns or sets if the font is formatted as small capital letters.
@property double spacing;  // Returns or sets the spacing in points between characters.
@property BOOL strikeThrough;  // Returns or sets if the font is formatted as strikethrough text.
@property BOOL subscript;  // Returns or sets if the font is formatted as subscript.
@property BOOL superscript;  // Returns or sets if the font is formatted as superscript.
@property MicrosoftWordE113 underline;  // Returns or sets the type of underline applied to the font.
@property (copy) NSColor *underlineColor;  // Returns or sets the RGB color of the underline for the font object.
@property MicrosoftWordMCoI underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.

- (void) growFont;  // Increases the font size to the next available size. If the selection or range contains more than one font size, each size is increased to the next available setting.
- (void) reset;  // Removes changes that were made to a font.
- (void) setAsFontTemplateDefault;  // Sets the font formatting as the default for the active document and all new documents based on the active template. The default font formatting is stored in the Normal style.
- (void) shrinkFont;  // Decreases the font size to the next available size. If the selection or range contains more than one font size, each size is decreased to the next available setting.

@end

// A representation of the options associated with footnotes.
@interface MicrosoftWordFootnoteOptions : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordTextRange *footnoteContinuationNotice;  // Returns a text range object that represents the footnote continuation notice.
@property (copy, readonly) MicrosoftWordTextRange *footnoteContinuationSeparator;  // Returns a text range object that represents the footnote continuation separator.
@property MicrosoftWordE174 footnoteLocation;  // Returns or sets the position of all footnotes.
@property MicrosoftWordE152 footnoteNumberStyle;  // Returns or sets the number style for footnotes.
@property MicrosoftWordE173 footnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. 
@property (copy, readonly) MicrosoftWordTextRange *footnoteSeparator;  // Returns a text range object that represents the footnote separator.
@property NSInteger footnoteStartingNumber;  // Returns or sets the starting footnote number. 

- (void) footnoteConvert;  // Converts endnotes to footnotes, or vice versa.
- (void) swapWithEndnotes;  // Converts all footnotes in a document to endnotes and vice versa.

@end

// Represents a footnote positioned at the bottom of the page or beneath text.
@interface MicrosoftWordFootnote : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) MicrosoftWordTextRange *noteReference;  // Returns a text range object that represents a footnote mark.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the footnote object.


@end

// Represents a single form field.
@interface MicrosoftWordFormField : MicrosoftWordBaseObject

@property BOOL calculateOnExit;  // Returns or sets if references to the specified form field are automatically updated whenever the field is exited.
@property (copy, readonly) MicrosoftWordCheckBox *checkBox;  // Returns the check box object associated with the form filed object.
@property (copy, readonly) MicrosoftWordDropDown *dropDown;  // Returns the drop down object associated with the form filed object.
@property BOOL enabled;  // Returns or sets if this form field object is enabled.
@property (copy) NSString *formFieldResult;  // Returns or sets a string that represents the result of the specified form field.
@property (readonly) MicrosoftWordE183 formFieldType;  // The type of this form field.
@property (copy) NSString *helpText;  // Returns or sets the text that's displayed in a message box when the form field has the focus and the user presses F1.
@property (copy) NSString *name;  // Returns or sets the name of the form field.
@property (copy, readonly) MicrosoftWordFormField *nextFormField;  // Returns the next form field object.
@property BOOL ownHelp;  // Returns or set the source of the text that's displayed in a message box when a form field has the focus and the user presses F1. If true, the text specified by the helptext property is displayed. If false, the text in the autotext entry is displayed
@property BOOL ownStatus;  // Returns or sets the source of the text that's displayed in the status bar when a form field has the focus. If true, the text specified by the status text property is displayed. If false, the text of the autotext entry is displayed.
@property (copy, readonly) MicrosoftWordFormField *previousFormField;  // Returns the previous form field object.
@property (copy) NSString *statusText;  // Returns or sets the text that's displayed in the status bar when a form field has the focus.
@property (copy, readonly) MicrosoftWordTextInput *textInput;  // Returns the text input object associated with the form filed object.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the form field object.


@end

// Represents a frame.
@interface MicrosoftWordFrame : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this frame object.
@property double height;  // Returns or sets the height of the object.
@property MicrosoftWordE134 heightRule;  // Returns or sets the rule for determining the height of the specified frame.
@property double horizontalDistanceFromText;  // Returns or sets the horizontal distance between a frame and the surrounding text, in points.
@property double horizontalPosition;  // Returns or sets the horizontal distance between the edge of the frame and the item specified by the relative horizontal position property.
@property BOOL lockAnchor;  // Returns or sets if the specified frame is locked. The frame anchor indicates where the frame will appear in Draft view. You cannot reposition a locked frame anchor.
@property MicrosoftWordE236 relativeHorizontalPosition;  // Returns or sets what the horizontal position of a frame is relative.
@property MicrosoftWordE237 relativeVerticalPosition;  // Returns or sets what the vertical position of a frame is relative.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the frame object.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the frame object.
@property BOOL textWrap;  // Returns or sets if document text wraps around the specified frame.
@property double verticalDistanceFromText;  // Returns or sets the vertical distance in points between a frame and the surrounding text.
@property double verticalPosition;  // Returns or sets the vertical distance between the edge of the frame and the item specified by the relative vertical position property. 
@property double width;  // Returns or sets the width of the object.
@property MicrosoftWordE134 widthRule;  // Returns or sets the rule used to determine the width of a frame.


@end

// Represents a single header or footer.
@interface MicrosoftWordHeaderFooter : MicrosoftWordBaseObject

- (SBElementArray *) pageNumbers;
- (SBElementArray *) shapes;

@property (readonly) MicrosoftWordE163 headerFooterIndex;  // Returns a constant that represents the specified header or footer in a document or section.
@property (readonly) BOOL isHeader;  // Returns true if this object is a header.
@property BOOL linkToPrevious;  // Returns or sets if the specified header or footer is linked to the corresponding header or footer in the previous section. When a header or footer is linked, its contents are the same as in the previous header or footer.
@property (copy, readonly) MicrosoftWordPageNumberOptions *pageNumberOptions;  // Return the page number options object associated with this header footer object.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the header or footer.


@end

// Represents a style used to build a table of contents or figures.
@interface MicrosoftWordHeadingStyle : MicrosoftWordBaseObject

@property NSInteger level;  // Returns or sets the level for the heading style in a table of contents or table of figures.
@property MicrosoftWordE184 style;  // Returns or sets the style associated with the heading style object.


@end

// Represents a hyperlink.
@interface MicrosoftWordHyperlinkObject : MicrosoftWordBaseObject

@property (copy) NSString *emailSubject;  // Returns or sets the text string for the specified hyperlink’s subject line. The subject line is appended to the hyperlink’s Internet address, or URL.
@property (readonly) BOOL extraInfoRequired;  // Returns true if extra information is required to resolve the specified hyperlink.
@property (copy) NSString *hyperlinkAddress;  // Returns or sets the address, for example, a file name or URL of the specified hyperlink. 
@property (readonly) MicrosoftWordMHlT hyperlinkType;  // The type of this hyperlink
@property (copy, readonly) NSString *name;  // The name of this hyperlink object.
@property (copy) NSString *screenTip;  // Returns or sets the text that appears as a screen tip when the mouse pointer is positioned over the specified hyperlink.
@property (copy, readonly) MicrosoftWordShape *shape;
@property (copy) NSString *subAddress;  // Returns or sets a named location in the destination of the specified hyperlink. 
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the hyperlink.
@property (copy) NSString *textToDisplay;  // Returns or sets the specified hyperlink's visible text in a document.

- (void) createNewDocumentForHyperlinkFileName:(NSString *)fileName editNow:(BOOL)editNow overwrite:(BOOL)overwrite;  // Creates a new document linked to the specified hyperlink.
- (void) followNewWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo;  // Displays a cached document associated with the specified hyperlink object, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

@end

// Represents a single index.
@interface MicrosoftWordIndex : MicrosoftWordBaseObject

@property BOOL accentedLetters;  // Returns or sets if the specified index contains separate headings for accented letters, for example, words that begin with À are under one heading and words that begin with A are under another.
@property MicrosoftWordE119 headingSeparator;  // Returns or sets the text between alphabetic groups, entries that start with the same letter in the index.  
@property MicrosoftWordE105 indexFilter;  // Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index. 
@property MicrosoftWordE214 indexType;  // Returns or sets the index type.
@property NSInteger numberOfColumns;  // Sets or returns the number of columns for each page of an index.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in an index. 
@property MicrosoftWordE106 sortBy;  // Returns or sets the sorting criteria for the specified index.
@property MicrosoftWordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in an index. 
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the index


@end

// Represents a custom key assignment in the current context.
@interface MicrosoftWordKeyBinding : MicrosoftWordBaseObject

@property (copy, readonly) SBObject *bindingContext;  // Returns an object that represents the storage location of the specified key binding. This property can return a document or template object.
@property (copy, readonly) NSString *bindingKeyString;  // Returns the key combination string for the specified keys.
@property (copy, readonly) NSString *command;  // Returns the command assigned to the specified key combination.
@property (copy, readonly) NSString *commandParameter;  // Returns the command parameter assigned to the specified shortcut key.
@property (readonly) MicrosoftWordE239 keyCategory;  // Returns the type of item assigned to the specified key binding.
@property (readonly) NSInteger keyCode;  // Returns a unique number for the first key in the specified key binding.  You create this number by using the build key code method.
@property (readonly) NSInteger key_code_2;  // Returns a unique number for the second key in the specified key binding. You create this number by using the build key code method.
- (BOOL) protected;  // Returns true if you cannot change the specified key binding in the customize keyboard dialog box. 

- (void) disable;  // Removes the specified key combination if it's currently assigned to a command. After you use this method, the key combination has no effect.
- (void) executeKeyBinding;  // Runs the command associated with the specified key combination.
- (void) rebindKeyCategory:(MicrosoftWordE239)keyCategory command:(NSString *)command commandParameter:(NSString *)commandParameter;  // Changes the command assigned to the specified key binding.

@end

// Represents the elements of a letter created by the letter wizard.
@interface MicrosoftWordLetterContent : MicrosoftWordBaseObject

@property (copy) NSString *attentionLine;  // Returns or sets the attention line text for a letter created by the letter wizard.
@property (copy) NSString *ccList;  // Returns or sets the carbon copy recipients for a letter created by the letter wizard.
@property (copy) NSString *closing;  // Returns or sets the closing text for a letter created by the letter wizard, for example, Sincerely yours.
@property (copy) NSString *dateFormat;  // Returns or sets the date for a letter created by the letter wizard. 
@property NSInteger enclosureCount;  // Returns or sets the number of enclosures for a letter created by the letter wizard. 
@property BOOL includeHeaderFooter;  // Returns or sets if the header and footer from the page design template are included in a letter created by the letter wizard.
@property MicrosoftWordE245 letterStyle;  // Returns or sets the layout of a letter created by the letter wizard.
@property BOOL letterhead;  // Returns or sets if space is reserved for a preprinted letterhead in a letter created by the letter wizard.
@property MicrosoftWordE246 letterheadLocation;  // Returns or sets the location of the preprinted letterhead in a letter created by the letter wizard.
@property double letterheadSize;  // Returns or sets the amount of space in points to be reserved for a preprinted letterhead in a letter created by the letter wizard.
@property (copy) NSString *mailingInstructions;  // Returns or sets the mailing instruction text for a letter created by the letter wizard, for example, Certified Mail.
@property (copy) NSString *pageDesign;  // Returns or sets the name of the template attached to the document created by the letter wizard.
@property (copy) NSString *recipientAddress;  // Returns or sets the return address for a letter created with the letter wizard.
@property (copy) NSString *recipientName;  // Returns or sets the name of the person who'll be receiving the letter created by the letter wizard.
@property (copy) NSString *recipientReference;  // Returns or sets the reference line, for example, In reply to: for a letter created by the letter wizard.
@property (copy) NSString *returnAddress;  // Returns or sets the return address for a letter created with the letter wizard. 
@property (copy) NSString *salutation;  // Returns or sets the salutation text for a letter created by the letter wizard.
@property MicrosoftWordE247 salutationType;  // Returns or sets the type of salutation for a letter created by the letter wizard.  
@property (copy, readonly) NSString *senderCity;
@property (copy) NSString *senderCompany;  // Returns or sets the company name of the person creating a letter with the letter wizard.
@property (copy) NSString *senderInitials;  // Returns or sets the initials of the person creating a letter with the letter wizard. 
@property (copy) NSString *senderJobTitle;  // Returns or sets the job title of the person creating a letter with the letter wizard.
@property (copy) NSString *senderName;  // Returns or sets the name of the person creating a letter with the letter wizard. 
@property (copy) NSString *subject;  // Returns or sets the subject text of a letter created by the letter wizard.


@end

// Represents line numbers in the left margin or to the left of each newspaper-style column.
@interface MicrosoftWordLineNumbering : MicrosoftWordBaseObject

@property BOOL activeLine;  // Returns or sets if line numbering is active for the specified document, section, or sections.
@property NSInteger countBy;  // Returns or sets the numeric increment for line numbers. For example, if the count by property is set to 5, every fifth line will display the line number. Line numbers are only displayed in print layout view and print preview.
@property double distanceFromText;  // Returns or sets the distance in points between the right edge of line numbers and the left edge of the document text.
@property MicrosoftWordE173 restartMode;  // Returns or sets the way line numbering runs— that is, whether it starts over at the beginning of a new page or section or runs continuously.
@property NSInteger startingNumber;  // Returns or sets the starting line number.


@end

// Represents the linking characteristics for a picture.
@interface MicrosoftWordLinkFormat : MicrosoftWordBaseObject

@property BOOL autoUpdate;  // Returns or sets if the specified link is updated automatically when the container file is opened or when the source file is changed.
@property (readonly) MicrosoftWordE200 linkType;  // Returns the link type.
@property BOOL locked;  // Returns or sets if inline shape object is locked to prevent automatic updating.
@property BOOL savePictureWithDocument;  // Returns or sets if the specified picture is saved with the document.
@property (copy) NSString *sourceFullName;  // Returns or sets the path and name of the source file for the specified picture.
@property (copy, readonly) NSString *sourceName;  // Returns the name of the source file for the specified picture.
@property (copy, readonly) NSString *sourcePath;  // Returns the path of the source file for the specified picture.

- (void) breakLink;  // Breaks the link between the source file and the specified picture.

@end

// Represents an item in a drop-down form field.
@interface MicrosoftWordListEntry : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of this list entry object.


@end

// Represents the list formatting attributes that can be applied to the paragraphs in a range.
@interface MicrosoftWordListFormat : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordWordList *WordList;  // Returns a Word list object that represents the first formatted list contained in the list format object.
@property NSInteger listLevelNumber;  // Returns or sets the list level for the first paragraph in the list format object.
@property (copy) MicrosoftWordInlineShape *listPictureBullet;
@property (copy, readonly) NSString *listString;  // Returns a string that represents the appearance of the list value of the first paragraph in the range for the list format object. For example, the second paragraph in an alphabetic list would return B.
@property (copy, readonly) MicrosoftWordListTemplate *listTemplate;  // Returns a list template object that represents the list formatting for the list format object.
@property (readonly) MicrosoftWordE159 listType;  // Returns the type of lists that are contained in the range for the list format object.
@property (readonly) NSInteger listValue;  // Returns the numeric value of the first paragraph in the range for the specified list format object. For example, the list value property applied to the second paragraph in an alphabetic list would return 2.
@property (readonly) BOOL singleList;  // Returns if the specified list format object contains only one list.
@property (readonly) BOOL singleListTemplate;  // True if the entire list format object uses the same list template

- (void) applyBulletDefaultDefaultListBehavior:(MicrosoftWordE289)defaultListBehavior;  // Adds bullets and formatting to the paragraphs in the range for the list format object. If the paragraphs are already formatted with bullets, this method removes the bullets and formatting.
- (void) applyListFormatTemplateListTemplate:(MicrosoftWordListTemplate *)listTemplate continuePreviousList:(BOOL)continuePreviousList applyTo:(MicrosoftWordE137)applyTo defaultListBehavior:(MicrosoftWordE289)defaultListBehavior;  // Applies a set of list-formatting characteristics to the list format object
- (void) applyNumberDefaultDefaultListBehavior:(MicrosoftWordE289)defaultListBehavior;  // Adds the default numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.
- (void) applyOutlineNumberDefaultDefaultListBehavior:(MicrosoftWordE289)defaultListBehavior;  // Adds the default outline-numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as an outline-numbered list, this method removes the numbers and formatting.
- (void) listIndent;  // Increases the list level of the paragraphs in the range for the list format object, in increments of one level.
- (void) listOutdent;  // Decreases the list level of the paragraphs in the range for the list format object, in increments of one level.

@end

// Represents a single gallery of list formats.
@interface MicrosoftWordListGallery : MicrosoftWordBaseObject

- (SBElementArray *) listTemplates;

- (BOOL) modifiedIndex:(NSInteger)index;  // if the specified list template is not the built-in list template for that position in the list gallery. Index goes from 1 to 7
- (void) resetListGalleryIndex:(NSInteger)index;  // Resets the list template specified by index for the specified list gallery to the built-in list template format.

@end

// Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list.
@interface MicrosoftWordListLevel : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) MicrosoftWordFont *fontObject;  // Returns the font object associated with this list level object.
@property (copy) NSString *linkedStyle;  // Returns or sets the name of the style that's linked to the specified list level object.
@property MicrosoftWordE143 listLevelAlignment;  // Returns or sets a constant that represents the alignment for the list level of the list template.
@property (copy) NSString *numberFormat;  // Returns or sets the number format for the specified list level.
@property double numberPosition;  // Returns or sets the position in points of the number or bullet for the specified list level object.
@property MicrosoftWordE151 numberStyle;  // Returns or sets the number style for the list level object.
@property (copy, readonly) MicrosoftWordInlineShape *pictureBullet;
@property NSInteger resetOnHigher;  // Returns or sets the list level that must appear before the specified list level restarts numbering at 1.
@property NSInteger startAt;  // Returns or sets the starting number for the specified list level object.
@property double tabPosition;  // Returns or sets the tab position for the specified list level object.
@property double textPosition;  // Returns or sets the position in points for the second line of wrapping text for the specified list level object.
@property MicrosoftWordE149 trailingCharacter;  // Returns or sets the character inserted after the number for the specified list level.

- (MicrosoftWordInlineShape *) applyPictureBulletPath:(NSString *)path;

@end

// Represents a single list template that includes all the formatting that defines a list.
@interface MicrosoftWordListTemplate : MicrosoftWordBaseObject

- (SBElementArray *) listLevels;

@property (copy) NSString *name;  // Returns or sets the name of this list template object.
@property BOOL outlineNumbered;  // Returns or sets if the specified list template object is outline numbered.

- (MicrosoftWordListTemplate *) convertLevel:(NSInteger)level;  // Converts a multiple-level list to a single-level list, or vice versa.

@end

// Represents a mailing label.
@interface MicrosoftWordMailingLabel : MicrosoftWordBaseObject

- (SBElementArray *) customLabels;

@property (copy) NSString *defaultLabelName;  // Returns or sets the name for the default mailing label.
@property MicrosoftWordE207 defaultLaserTray;  // Returns or sets the default paper tray that contains sheets of mailing labels
@property BOOL defaultPrintBarCode;  // Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set

- (MicrosoftWordDocument *) createNewMailingLabelDocumentName:(NSString *)name address:(NSString *)address autoText:(NSString *)autoText extractAddress:(BOOL)extractAddress laserTray:(MicrosoftWordE207)laserTray singleLabel:(BOOL)singleLabel row:(NSInteger)row column:(NSInteger)column;  // Creates a new label document using either the default label options or ones that you specify. Returns a document object that represents the new document.
- (void) printOutMailingLabelName:(NSString *)name address:(NSString *)address extractAddress:(BOOL)extractAddress laserTray:(MicrosoftWordE207)laserTray singleLabel:(BOOL)singleLabel row:(NSInteger)row column:(NSInteger)column;  // Prints a label or a page of labels with the same address.

@end

// Represents an equation that has an accent mark above the base.
@interface MicrosoftWordMathAccent : MicrosoftWordBaseObject

@property NSInteger char1;  // Gets or sets an integer that represents the accent character for the accent object.
@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.


@end

// Represents an individual entry in the Math AutoCorrect engine.
@interface MicrosoftWordMathAutocorrectEntry : MicrosoftWordBaseObject

@property (copy) NSString *autocorrectValue;  // Gets or sets a string that represents the contents of an equation auto correct entry.
@property (readonly) NSInteger entry_index;  // Gets an integer that represents the position of an item in the collection.
@property (copy) NSString *name;  // Gets or sets a string that represents the name of an equation auto correct entry.


@end

// Represents the Math AutoCorrect feature in Microsoft Word.
@interface MicrosoftWordMathAutocorrect : MicrosoftWordBaseObject

- (SBElementArray *) mathAutocorrectEntries;
- (SBElementArray *) mathRecognizedFunctions;

@property BOOL replaceText;  // Gets or sets whether Microsoft Word automatically replaces strings in equations with the corresponding math AutoCorrect definitions.
@property BOOL useOutsideEquations;  // Gets or sets whether Microsoft Word uses math autocorrect rules outside equations in a document.

- (void) addMathAcEntryName:(NSString *)name value:(NSString *)value;  // Creates an equation auto correct entry.
- (void) addMathRecognizedFunctionName:(NSString *)name;  // Creates a new recognized function.

@end

// Represents the mathematical overbar for an object in an equation.
@interface MicrosoftWordMathBar : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isTopBar;  // Gets or sets the position of a bar in a bar object. True specifies a mathematical overbar. False specifies a mathematical underbar


@end

// Represents an invisible box around an equation or part of an equation to which you can assign properties that affect the layout or mathematical formatting of the entire box.
@interface MicrosoftWordMathBorderBox : MicrosoftWordBaseObject

@property BOOL bottomLeftToTopRightStrikethrough;  // Gets or sets if a diagonal strikethrough from lower left to upper right is drawn.
@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hideBottom;  // Gets or sets whether to hide the bottom border of an equation's bounding box.
@property BOOL hideLeft;  // Gets or sets whether to hide the left border of an equation's bounding box.
@property BOOL hideRight;  // Gets or sets whether to hide the right border of an equation's bounding box.
@property BOOL hideTop;  // Gets or sets whether to hide the top border of an equation's bounding box.
@property BOOL horizontalStrikethrough;  // Gets or sets if a horizontal strikethrough is drawn.
@property BOOL topLeftToBottomRightStrikethrough;  // Gets or sets if a diagonal strikethrough from upper left to lower right is drawn.
@property BOOL verticalStrikethrough;  // Gets or sets if vertical strikethrough is drawn.


@end

// Represents an invisible box around an equation or part of an equation to which you can apply properties that affect the mathematical or formatting properties, such as line breaks.
@interface MicrosoftWordMathBox : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isDifferential;  // Gets or sets whether the box acts as the mathematical differential, in which case the box receives the appropriate horizontal spacing for a differential.
@property BOOL isOperatorEmulator;  // Gets or sets if the box and its contents behave as a single operator and inherits the properties of an operator.
@property BOOL nonBreaking;  // Gets or sets whether breaks are allowed inside the box object.


@end

// Represents individual line breaks in an equation.
@interface MicrosoftWordMathBreak : MicrosoftWordBaseObject

@property NSInteger alignAt;  // Gets or sets an integer that represents the operator in one line, to which to align consecutive lines in an equation.
@property (copy, readonly) MicrosoftWordTextRange *textRange;  // Returns a text range object that represents the portion of a document that is contained in the specified object.


@end

// Represents a delimiter object, consisting of opening and closing delimiters, e.g. parentheses, braces, brackets, or vertical bars, and one or more elements contained inside the delimiters.
@interface MicrosoftWordMathDelimiter : MicrosoftWordBaseObject

- (SBElementArray *) mathObjects;

@property NSInteger beginningCharacter;  // Gets or sets an Integer that represents the beginning delimiter character in a delimiter object.
@property MicrosoftWordE325 delimiterShape;  // Gets or sets the appearance of delimiters, e.g. parentheses, braces, and brackets, in relationship to the content that they surround.
@property NSInteger endingCharacter;  // Gets or sets an Integer that represents the ending delimiter character in a delimiter object.
@property BOOL hideLeftDelimiter;  // Gets or sets whether to hide the opening delimiter in a delimiter object.
@property BOOL hideRightDelimiter;  // Gets or sets whether to hide the closing delimiter in a delimiter object.
@property NSInteger separatorCharacter;  // Gets or sets an Integer that represents the separator character in a delimiter object when the delimiter object contains two or more arguments.
@property BOOL shouldGrow;  // Gets or sets whether delimiter characters grow to the full height of the arguments that they contain.


@end

// Represents a mathematical equation array object, consisting of one or more equations that can be vertically justified as a unit respect to surrounding text on the line.
@interface MicrosoftWordMathEquationArray : MicrosoftWordBaseObject

- (SBElementArray *) mathObjects;

@property BOOL distributeEqually;  // Gets or sets if the equations in an equation array are distributed equally within the margins of its container, such as a column, cell, or page width.
@property NSInteger rowSpacing;  // Gets or sets an integer that represents the spacing between the rows in an equation array.
@property MicrosoftWordE323 rowSpacingRule;  // Gets or sets the spacing rule that defines spacing in an equation array.
@property BOOL useMaxWidth;  // Gets or sets whether the equations in an equation array are spaced to the maximum width of the equation array.
@property MicrosoftWordE321 verticalAlignment;  // Gets or sets the type of vertical alignment for an equation array with respect to the text that surrounds the array.


@end

// Represents a fraction, consisting of a numerator and denominator separated by a fraction bar. The fraction bar can be horizontal or diagonal, depending on the fraction properties.
@interface MicrosoftWordMathFraction : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *denominator;  // Returns a math object object that represents the denominator for an equation that contains a fraction.
@property MicrosoftWordE322 fractionType;  // Gets or sets the layout of a fraction, whether it is stacked, skewed, linear, or without a fraction bar.
@property (copy, readonly) MicrosoftWordMathObject *numerator;  // Returns a math object object that represents the numerator for the fraction.


@end

// Represents a math func object that represents a type of mathematical function that consists of a function name, such as sin or cos, and an argument.
@interface MicrosoftWordMathFunc : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *funcName;  // Returns an OMath object that represents the name of a mathematical function, such as sin or cos.


@end

// Represents a mathematical function or structure such as fractions, integrals, sums, and radicals.
@interface MicrosoftWordMathFunction : MicrosoftWordBaseObject

- (SBElementArray *) mathObjects;

@property (copy, readonly) MicrosoftWordMathAccent *accent;  // Returns a math accent object that represents a base character with a combining accent mark.
@property (copy, readonly) MicrosoftWordMathBorderBox *borderBox;  // Returns a math border box object that represents a border drawn around an equation or part of an equation.
@property (copy, readonly) MicrosoftWordMathBox *box;  // Returns a math box object to which you can apply properties.
@property (copy, readonly) MicrosoftWordMathDelimiter *delimiter;  // Returns an math delimiter object that represents the delimiter function
@property (copy, readonly) MicrosoftWordMathEquationArray *equationArray;  // Returns a math equation array object that represents an equation array function.
@property (copy, readonly) MicrosoftWordMathFraction *fraction;  // Returns a math fraction object that represents a fraction.
@property (copy, readonly) MicrosoftWordMathFunc *func;  // Returns a math func object that represents a type of mathematical function.
@property (readonly) MicrosoftWordE319 functionType;  // Gets the type of the function.
@property (copy, readonly) MicrosoftWordMathGroupChar *groupChar;  // Returns a math group char object that represents a horizontal character placed above or below text in an equation.
@property (copy, readonly) MicrosoftWordMathLeftScripts *leftScripts;
@property (copy, readonly) MicrosoftWordMathLowerLimit *lowerLimit;  // Returns a math lower limit object that represents the lower limit for a function.
@property (copy, readonly) MicrosoftWordMathObject *mathObj;  // Returns an OMath object that represents the equation.
@property (copy, readonly) MicrosoftWordMathMatrix *matrix;  // Returns a math matrix object that represents a mathematical matrix.
@property (copy, readonly) MicrosoftWordMathNary *nary;  // Returns a math nary object that represents the n-ary operation.
@property (copy, readonly) MicrosoftWordMathBar *overbar;  // Returns a math bar object that represents the mathematical overbar for an object.
@property (copy, readonly) MicrosoftWordMathPhantom *phantom;  // Returns a math phantom object that represents an object used for advanced layout of an equation.
@property (copy, readonly) MicrosoftWordMathRadical *radical;  // Returns a math radical object that represents the mathematical radical function.
@property (copy, readonly) MicrosoftWordMathSubAndSuperScript *subAndSuperScript;  // Returns a math sub and super script object that represents a mathematical subscript-superscript object that consists of a base, a subscript, and a superscript.
@property (copy, readonly) MicrosoftWordMathSubscript *subscript;  // Returns a math subscript object that represents the mathematical subscript function.
@property (copy, readonly) MicrosoftWordMathSuperscript *superscript;  // Returns a math superscript object that represents the mathematical superscript function.
@property (copy, readonly) MicrosoftWordTextRange *textRange;  // Returns a text range object that represents the portion of a document that is contained in the math function.
@property (copy, readonly) MicrosoftWordMathUpperLimit *upperLimit;  // Returns a math upper limit object that represents upper limit function.


@end

// Represents a group character object, consisting of a character drawn above or below text, often with the purpose of visually grouping items.
@interface MicrosoftWordMathGroupChar : MicrosoftWordBaseObject

@property BOOL alignTop;  // Returns or sets whether the grouping character is aligned vertically with the surrounding text or whether the base text that is either above or below the grouping character is aligned vertically with the surrounding text.
@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isOnTop;  // Gets or sets whether the grouping character is placed above the base text of the group character object. False displays the group character under the base text.
@property NSInteger theCharacter;  // Gets or sets an integer that represents the character placed above or below text in a group character object.


@end

// Represents an equation that contains a superscript or subscript to the left of the base.
@interface MicrosoftWordMathLeftScripts : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *subscript;  // Returns a math object object that represents the subscript for a pre-sub-superscript object.
@property (copy, readonly) MicrosoftWordMathObject *superscript;  // Returns a math object object that represents the superscript for a pre-sub-superscript object.

- (MicrosoftWordMathFunction *) convertLeftScriptsToSubAndSuperScripts;  // Converts an equation with a superscript or subscript to the left of the base of the equation to an equation with a base of a superscript or subscript.

@end

// Represents the lower limit mathematical construct, consisting of text on the baseline and reduced-size text immediately below it.
@interface MicrosoftWordMathLowerLimit : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *limit;  // Returns a math object object that represents the limit of the lower limit object. 

- (MicrosoftWordMathFunction *) lowerLimitToUpperLimit;

@end

// Represents a matrix column.
@interface MicrosoftWordMathMatrixColumn : MicrosoftWordBaseObject

- (SBElementArray *) mathObjects;

@property NSInteger columnIndex;  // Gets or sets an integer that represents the ordinal position of a matrix column within the collection of matrix columns.
@property MicrosoftWordE320 horizontalAlignment;  // Gets or sets the horizontal alignment for arguments in a matrix column.


@end

// Represents a matrix row.
@interface MicrosoftWordMathMatrixRow : MicrosoftWordBaseObject

- (SBElementArray *) mathObjects;

@property NSInteger rowIndex;  // Gets or sets an integer that represents the ordinal position of a matrix row within the collection of matrix rows.


@end

// Represents an equation matrix.
@interface MicrosoftWordMathMatrix : MicrosoftWordBaseObject

- (SBElementArray *) mathMatrixRows;
- (SBElementArray *) mathMatrixColumns;

@property NSInteger columnGap;  // Gets or sets an integer that represents the spacing between columns in a matrix.
@property MicrosoftWordE323 columnGapRule;  // Gets or sets the spacing rule for the space that appears between columns in a matrix.
@property NSInteger columnSpacing;  // Gets or sets an integer that represents the spacing for columns in a matrix.
@property BOOL hidePlaceholders;  // Gets or sets whether placeholders in a matrix are hidden from display.
@property NSInteger rowSpacing;  // Gets or sets an integer that represents the spacing for rows in a matrix. 
@property MicrosoftWordE323 rowSpacingRule;  // Gets or sets the spacing rule for rows in a matrix.
@property MicrosoftWordE321 verticalAlignment;  // Gets or sets the vertical alignment for a matrix.

- (MicrosoftWordMathMatrixColumn *) addMatrixColumnBeforeColumn:(MicrosoftWordMathMatrixColumn *)beforeColumn;  // Creates a matrix column and adds it to a matrix and returns a math matrix column object.
- (MicrosoftWordMathMatrixRow *) addMatrixRowBeforeRow:(MicrosoftWordMathMatrixRow *)beforeRow;  // Creates a matrix row and adds it to a matrix and returns a math matrix row object.

@end

// Represents the mathematical n-ary object, consisting of an n-ary object, a base/operand, and optional upper limits and lower limits.
@interface MicrosoftWordMathNary : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hidesLowerLimit;  // Gets or sets whether to hide the lower limit of an n-ary operator.
@property BOOL hidesUpperLimit;  // Gets or sets whether to hide the upper limit of an n-ary operator.
@property (copy, readonly) MicrosoftWordMathObject *lowerLimit;  // Returns a math object object that represents the lower limit of an n-ary operator.
@property NSInteger operatorCharacter;  // Gets or sets an integer that represents a character used as the n-ary operator.
@property BOOL shouldGrow;  // Gets or sets whether n-ary operators grow to the full height of the arguments that they contain.
@property (copy, readonly) MicrosoftWordMathObject *upperLimit;  // Returns a math object object that represents the upper limit of an n-ary operator.
@property BOOL useSubAndSuperScriptPositioning;  // Gets or sets the positioning of n-ary limits in the subscript-superscript or upper limit-lower limit position.


@end

// Represents an equation
@interface MicrosoftWordMathObject : MicrosoftWordBaseObject

- (SBElementArray *) mathFunctions;
- (SBElementArray *) mathBreaks;

@property NSInteger alignPoint;  // Gets or sets an integer that represents the character position of the alignment point in the equation.
@property (readonly) NSInteger argumentIndex;  // Gets an integer that specifies the argument index of this component relative to the containing math object.
@property (copy, readonly) MicrosoftWordMathFunction *containingFunction;  // Gets the math function that represents the parent, or containing, function.
@property MicrosoftWordE324 displayType;  // Sets or returns the display type of the equation.
@property MicrosoftWordE326 justification;  // Gets or sets the justification for the equation.
@property (readonly) NSInteger nestingLevel;  // Returns an integer that represents the nesting level for the math object.
@property (copy, readonly) MicrosoftWordMathObject *parentArgument;  // Gets a math object that represents the parent, or containing, argument.
@property (copy, readonly) MicrosoftWordMathMatrixColumn *parentColumn;  // Gets a math matrix column object that represents the parent column in a matrix.
@property (copy, readonly) MicrosoftWordMathObject *parentObject;  // Gets the math object that represents the parent object.
@property (copy, readonly) MicrosoftWordMathMatrixRow *parentRow;  // Gets a math matrix row object that represents the parent row in a matrix.
@property NSInteger scriptSize;  // Gets or sets an integer that represents the script size of an argument, for example, text, script, or script-script.
@property (copy, readonly) MicrosoftWordTextRange *textRange;  // Gets a text range object that represents the portion of a document that contains the equation.

- (void) addMathFunctionInLocation:(MicrosoftWordTextRange *)inLocation mathFunctionType:(MicrosoftWordE319)mathFunctionType numberOfArguments:(NSInteger)numberOfArguments numberOfColumns:(NSInteger)numberOfColumns;  // Inserts a new structure, such as a fraction, into an equation at the specified position.
- (void) buildUp;  // Converts an equation to professional/built up format.
- (void) convertToLiteralText;  // Converts an equation to literal text.
- (void) convertToMathText;  // Converts an equation to math text.
- (void) convertToNormalText;  // Converts an equation to normal text.
- (MicrosoftWordMathBreak *) insertMathBreakAtRange:(MicrosoftWordTextRange *)atRange;  // Inserts a break into an equation and returns a math break object that represents the break.
- (void) linearize;  // Converts an equation to linear/built down format.

@end

// Represents a phantom object, which has two primary uses: 1. adding the spacing of the phantom base without displaying that base or 2. suppressing part of the glyph from spacing considerations.
@interface MicrosoftWordMathPhantom : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL showsInvisibles;  // Gets or sets whether the contents of a phantom object are visible.
@property BOOL smash;  // Gets or sets if the contents of the phantom are visible but that the height is not taken into account in the spacing of the layout.
@property BOOL transparent;  // Gets or sets whether a phantom object is transparent.
@property BOOL zeroAscent;  // Gets or sets whether the ascent of the phantom contents is ignored in the spacing of the layout.
@property BOOL zeroDescent;  // Gets or sets whether the descent of the phantom contents is ignored in the spacing of the layout.
@property BOOL zeroWidth;  // Gets or sets whether the width of a phantom object is ignored in the spacing of the layout.


@end

// Represents the mathematical radical object, consisting of a radical, a base, and an optional degree.
@interface MicrosoftWordMathRadical : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *degree;  // Returns a math object object that represents the degree for a radical.
@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hideDegree;  // Gets or sets whether to hide the degree for the radical.


@end

@interface MicrosoftWordMathRecognizedFunction : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Gets an integer that represents the position of an item in the collection.
@property (copy, readonly) NSString *name;  // Gets a string that represents the name of an equation recognized function.


@end

// Represents an equation with a base that contains a superscript or subscript.
@interface MicrosoftWordMathSubAndSuperScript : MicrosoftWordBaseObject

@property BOOL alignScripts;  // Gets or sets whether to horizontally align subscripts and superscripts in the sub-superscript object.
@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *subscript;  // Returns a math object object that represents the subscript for a subscript-superscript object.
@property (copy, readonly) MicrosoftWordMathObject *superscript;  // Returns a math object object that represents the superscript for a subscript-superscript object.

- (MicrosoftWordMathFunction *) convertSubAndSuperScriptsToLeftScripts;  // Converts an equation with a base superscript or subscript to an equation with a superscript or subscript to the left of the base.
- (MicrosoftWordMathFunction *) removeSubscript;  // Removes the subscript for an equation and returns a math function object that represents the updated equation without the subscript.
- (MicrosoftWordMathFunction *) removeSuperscript;  // Removes the superscript for an equation and returns a math function object that represents the updated equation without the superscript.

@end

// Represents an equation with a base that contains a subscript.
@interface MicrosoftWordMathSubscript : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *subscript;  // Returns a math object that represents the subscript for a subscript object.


@end

// Represents an equation with a base that contains a superscript.
@interface MicrosoftWordMathSuperscript : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *superscript;  // Returns a math object object that represents the superscript for a superscript object.


@end

// Represents the upper limit mathematical construct, consisting of text on the baseline and reduced-size text immediately above it.
@interface MicrosoftWordMathUpperLimit : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) MicrosoftWordMathObject *limit;  // Returns a math object object that represents the limit of the upper limit object.

- (MicrosoftWordMathFunction *) upperLimitToLowerLimit;

@end

// Represents options associated with page number objects
@interface MicrosoftWordPageNumberOptions : MicrosoftWordBaseObject

@property MicrosoftWordE120 chapterPageSeparator;  // Returns or sets the separator character used between the chapter number and the page number. 
@property NSInteger headingLevelForChapter;  // Returns or sets the heading level style that's applied to the chapter titles in the document. Can be a number from zero through 8, corresponding to heading levels 1 through 9.
@property BOOL includeChapterNumber;  // Returns or sets if a chapter number is included with page numbers.
@property MicrosoftWordE154 numberStyle;  // Returns or sets the number style for the page number objects.
@property BOOL restartNumberingAtSection;  // Returns or sets if page numbering starts at 1 again at the beginning of the specified section.
@property BOOL showFirstPageNumber;  // Returns or sets if the page number appears on the first page in the section.
@property NSInteger startingNumber;  // Returns or sets the starting page number.


@end

// Represents a page number in a header or footer.
@interface MicrosoftWordPageNumber : MicrosoftWordBaseObject

@property MicrosoftWordE121 alignment;  // Returns or sets a constant that represents the alignment for the page number.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.


@end

// Represents the page setup description.
@interface MicrosoftWordPageSetup : MicrosoftWordBaseObject

- (SBElementArray *) textColumns;

@property double bottomMargin;  // Returns or sets the distance in points between the bottom edge of the page and the bottom boundary of the body text.
@property double charsLine;  // Returns or sets the number of characters per line in the document grid.
@property BOOL differentFirstPageHeaderFooter;  // Returns or set if a different header or footer is used on the first page.
@property MicrosoftWordE207 firstPageTray;  // Returns or sets the paper tray to use for the first page of a document or section.
@property double footerDistance;  // Returns or sets the distance in points between the footer and the bottom of the page.
@property double gutter;  // Returns or sets the amount in points of extra margin space added to each page in a document or section for binding. 
@property MicrosoftWordE283 gutterPosition;  // Returns or sets on which side the gutter appears in a document.
@property double headerDistance;  // Returns or sets the distance in points between the header and the top of the page.
@property MicrosoftWordE278 layoutMode;  // Returns or sets the layout mode for the current document.
@property double leftMargin;  // Returns or sets the distance in points between the left edge of the page and the left boundary of the body text.
@property BOOL lineBetweenTextColumns;  // Returns or sets if vertical lines appear between all the columns.
@property (copy) MicrosoftWordLineNumbering *lineNumbering;  // Returns or sets the line numbering object that represents the line numbers for the specified page setup object.
@property double linesPage;  // Returns or sets the number of lines per page in the document grid.
@property BOOL mirrorMargins;  // Returns or sets if the inside and outside margins of facing pages are the same width.
@property BOOL oddAndEvenPagesHeaderFooter;  // Returns or sets if the specified page setup object has different headers and footers for odd-numbered and even-numbered pages.
@property MicrosoftWordE208 orientation;  // Returns or sets the orientation of the page.
@property MicrosoftWordE207 otherPagesTray;  // Returns or sets the paper tray to be used for all but the first page of a document or section.
@property double pageHeight;  // Returns or sets the height of the page in points.
@property double pageWidth;  // Returns or sets the width of the page in points.
@property MicrosoftWordE232 paperSize;  // Returns or sets the paper size.
@property double rightMargin;  // Returns or sets the distance in points between the right edge of the page and the right boundary of the body text.
@property MicrosoftWordE219 sectionStart;  // Returns or sets the type of section break for the specified object.
@property BOOL showGrid;  // Determines whether to show the grid.
@property double spacingBetweenTextColumns;  // Returns or sets the spacing in points between columns.
@property BOOL suppressEndnotes;  // Returns or sets if endnotes are printed at the end of the next section that doesn't suppress endnotes. Suppressed endnotes are printed before the endnotes in that section.
@property BOOL textColumnsEvenlySpaced;  // Returns or sets if text columns are evenly spaced.
@property double topMargin;  // Returns or sets the distance in points between the top edge of the page and the top boundary of the body text.
@property MicrosoftWordE146 verticalAlignment;  // Returns or sets the vertical alignment of text on each page in a document or section.
@property double widthOfTextColumns;  // Returns or sets the width of all text columns

- (void) setAsPageSetupTemplateDefault;  // Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.
- (void) setNumberOfTextColumnsNumberOfColumns:(NSInteger)numberOfColumns;  // Arranges text into the specified number of text columns. 
- (void) togglePortrait;  // Switches between portrait and landscape page orientations for a document or section.

@end

// Represents a window pane.
@interface MicrosoftWordPane : MicrosoftWordBaseObject

@property BOOL browseToWindow;  // Returns or sets if lines wrap at the right edge of the pane rather than at the right margin of the page.
@property (readonly) NSInteger browseWidth;  // Returns the width in points of the area in which text wraps in the specified pane. This property works only when you're in web layout view.
@property BOOL displayRulers;  // Returns or sets if rulers are displayed for the window
@property BOOL displayVerticalRuler;  // Returns or sets if vertical rulers are displayed for the window
@property (copy, readonly) MicrosoftWordDocument *document;  // Returns a document object associated with this pane.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property NSInteger horizontalPercentScrolled;  // Returns or sets the horizontal scroll position as a percentage of the pane width.
@property NSInteger minimumFontSize;  // Returns or sets the minimum font size in points displayed for the specified pane.
@property (copy, readonly) MicrosoftWordPane *nextPane;  // Returns the next pane object.
@property (copy, readonly) MicrosoftWordPane *previousPane;  // Returns the previous pane object.
@property (copy, readonly) MicrosoftWordSelectionObject *selection;  // Returns the selection object that represents a selected text range or the insertion point.
@property NSInteger verticalPercentScrolled;  // Returns or sets the vertical scroll position as a percentage of the pane width.
@property (copy, readonly) MicrosoftWordView *view;  // Returns a view object that represents the view for the pane.

- (MicrosoftWordZoom *) getZoomZoomType:(MicrosoftWordE202)zoomType;  // Returns a zoom object of the specified type for this pane.

@end

@interface MicrosoftWordRangeEndnoteOptions : MicrosoftWordBaseObject

@property MicrosoftWordE175 endnoteLocation;  // Returns or sets the position of endnotes in a range or selection.
@property MicrosoftWordE152 endnoteNumberStyle;  // Returns or sets the number style for endnotes in a range or selection.
@property MicrosoftWordE173 endnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks.
@property NSInteger endnoteStartingNumber;  // Returns or sets the starting endnote number in a range or selection.


@end

@interface MicrosoftWordRangeFootnoteOptions : MicrosoftWordBaseObject

@property MicrosoftWordE174 footnoteLocation;  // Returns or sets the position of footnotes in a range or selection.
@property MicrosoftWordE152 footnoteNumberStyle;  // Returns or sets the number style for footnotes in a range or selection.
@property MicrosoftWordE173 footnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. 
@property NSInteger footnoteStartingNumber;  // Returns or sets the starting number for footnotes in a range or selection.


@end

// Represents a recently used file.
@interface MicrosoftWordRecentFile : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the recently used file.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the recent file.
@property BOOL readOnly;  // Returns or sets if changes to the document cannot be saved to the original document. 

- (MicrosoftWordDocument *) openRecentFile;  // Opens the recent file and returns a document object.

@end

// Represents the replace criteria for a find-and-replace operation.
@interface MicrosoftWordReplacement : MicrosoftWordBaseObject

@property (copy) NSString *content;  // Returns or sets the text to replace in the specified range or selection.
@property (copy, readonly) MicrosoftWordFont *fontObject;  // The font object associated with this replacement object.
@property (copy, readonly) MicrosoftWordFrame *frame;  // The frame object associated with this replacement object.
@property NSInteger highlight;  // Returns or sets if highlight formatting is applied to the replacement text.
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the replacement object
@property MicrosoftWordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;  // Returns or set the paragraph format object associated with this replacement object.
@property MicrosoftWordE184 style;  // Returns or sets the Word style associated with the replacement object.


@end

// Property of View: a person who has made tracked changes in the viewed document.
@interface MicrosoftWordReviewer : MicrosoftWordBaseObject

@property BOOL visible;


@end

// Represents a change marked with a revision mark.
@interface MicrosoftWordRevision : MicrosoftWordBaseObject

@property (copy, readonly) NSString *author;  // Returns the name of the user who made the specified tracked change. 
@property (copy, readonly) id cells;
@property (copy, readonly) NSDate *dateValue;  // The date and time that the tracked change was made.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *formatDescription;
@property (copy, readonly) MicrosoftWordTextRange *movedRange;
@property (readonly) MicrosoftWordE216 revisionType;  // Returns the revision type.
@property (copy, readonly) id style;
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the revision


@end

// Represents the current selection in a window or pane. A selection represents either a selected or highlighted area in the document, or it represents the insertion point if nothing in the document is selected.
@interface MicrosoftWordSelectionObject : MicrosoftWordBaseObject

- (SBElementArray *) tables;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) characters;
- (SBElementArray *) footnotes;
- (SBElementArray *) endnotes;
- (SBElementArray *) WordComments;
- (SBElementArray *) cells;
- (SBElementArray *) sections;
- (SBElementArray *) paragraphs;
- (SBElementArray *) fields;
- (SBElementArray *) formFields;
- (SBElementArray *) frames;
- (SBElementArray *) bookmarks;
- (SBElementArray *) hyperlinkObjects;
- (SBElementArray *) columns;
- (SBElementArray *) rows;
- (SBElementArray *) inlineShapes;
- (SBElementArray *) shapes;
- (SBElementArray *) mathObjects;

@property (readonly) BOOL IPAtEndOfLine;  // Returns true if the insertion point is at the end of a line that wraps to the next line. False if the selection isn't collapsed, if the insertion point isn't at the end of a line, or if the insertion point is positioned before a paragraph mark.
@property (readonly) NSInteger bookmarkId;  // Returns the number of the bookmark that encloses the beginning of the selection. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with the selection
@property (copy, readonly) MicrosoftWordColumnOptions *columnOptions;  // Returns a column options object for this selection.
@property BOOL columnSelectMode;  // Returns or set if column selection mode is active. When this mode is active, the letters COL appear on the status bar.
@property (copy) NSString *content;  // Returns or sets the text in the selection.
@property (copy, readonly) MicrosoftWordDocument *document;  // Returns the document object associated with the selection.
@property (copy, readonly) MicrosoftWordEndnoteOptions *endnoteOptions;  // Returns the endnote options object for the selection.
@property BOOL extendMode;  // Returns or set if extend mode is active.
@property (copy, readonly) MicrosoftWordFind *findObject;  // Returns the find object associated with the selection.
@property double fitTextWidth;  // Returns or sets the width in the current measurement units in which Microsoft Word fits the text in the current selection. 
@property (copy, readonly) MicrosoftWordFont *fontObject;  // The font object associated with the selection.
@property (copy, readonly) MicrosoftWordFootnoteOptions *footnoteOptions;  // Returns the footnote options object for the selection.
@property (copy) MicrosoftWordTextRange *formattedText;  // Returns or sets a text range object that includes the formatted text in the selection.
@property (copy, readonly) MicrosoftWordHeaderFooter *headerFooterObject;  // Returns the header footer object associated with the selection.
@property (readonly) BOOL isEndOfRowMark;  // Returns true if the selection is collapsed and is located at the end-of-row mark in a table.
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the selection object
@property MicrosoftWordE182 languageIDEastAsian;  // Returns or sets an east asian language for the selection.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property MicrosoftWordE270 orientation;  // Returns or sets the orientation of text in the selection when the text direction feature is enabled.
@property (copy) MicrosoftWordPageSetup *pageSetup;  // Returns or set the page setup object associated with the selection.
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;  // Returns or set the paragraph object associated with the selection.
@property (readonly) NSInteger previousBookmarkId;  // Returns the number of the last bookmark that starts before or at the same place as the selection, It returns zero if there's no corresponding bookmark.
@property (copy, readonly) MicrosoftWordRangeEndnoteOptions *rangeEndnoteOptions;
@property (copy, readonly) MicrosoftWordRangeFootnoteOptions *rangeFootnoteOptions;
@property (copy, readonly) MicrosoftWordRowOptions *rowOptions;
@property NSInteger selectionEnd;  // Returns or sets the ending character position of the selection.
@property MicrosoftWordE261 selectionFlags;  // Returns or sets properties of the selection.
@property (readonly) BOOL selectionIsActive;  // Returns if the selection is active.
@property NSInteger selectionStart;  // Returns or sets the starting character position of the selection.
@property (readonly) MicrosoftWordE209 selectionType;  // Returns the selection type.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the selection
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are included in the elements of the selection
@property BOOL startIsActive;  // Returns or sets if the beginning of the selection is active. If the selection is not collapsed to an insertion point, either the beginning or the end of the selection is active.
@property (readonly) NSInteger storyLength;  // Returns the number of characters in the story that contains the selection
@property (readonly) MicrosoftWordE160 storyType;  // Returns the story type for the selection.
@property MicrosoftWordE184 style;  // Returns or sets the Word style associated with the selection object.
@property MicrosoftWordE182 supplementalLanguageID;  // Returns or sets the language for the selection
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the selection

- (double) calculateSelection;  // Calculates a mathematical expression within a selection.
- (void) copyFormat;//NS_RETURNS_NOT_RETAINED;  // Copies the character formatting of the first character in the selected text. If a paragraph mark is selected, Word copies paragraph formatting in addition to character formatting.
- (void) createTextbox;  // Adds a default-size text box around the selection. If the selection is an insertion point, this method changes the pointer to a cross-hair pointer so that the user can draw a text box.
- (MicrosoftWordTextRange *) endKeyMove:(MicrosoftWordE295)move extend:(MicrosoftWordE249)extend;  // Moves or extends the selection to the end of the specified unit. This method returns the new range object or missing value if an error occurred.
- (void) escapeKey;  // Cancels a mode such as extend or column select.
- (id) expandBy:(MicrosoftWordE129)by;  // Expands the selection. Returns the number of characters added to the range or selection.
- (void) extendCharacter:(NSString *)character;  // Extends the selection to the next larger unit of text. The progression of selected units of text is as follows: word, sentence, paragraph, section, entire document. The selection is extended by moving the active end of the selection.
- (MicrosoftWordField *) getNextField;  // Returns the next field object.
- (MicrosoftWordField *) getPreviousField;  // Returns the previous field object.
- (NSString *) getSelectionInformationInformationType:(MicrosoftWordE266)informationType;  // Returns information about the selection. 
- (NSInteger) homeKeyMove:(MicrosoftWordE295)move extend:(MicrosoftWordE249)extend;  // Moves or extends the selection to the beginning of the specified unit. This method returns the new range object or missing value if an error occurred.
- (void) insertCellsShiftCells:(MicrosoftWordE135)shiftCells;  // Adds cells to an existing table. The number of cells inserted is equal to the number of cells in the selection.
- (void) insertColumnsPosition:(MicrosoftWordEiRL)position;  // Inserts columns to the left of the column that contains the selection. If the selection isn't in a table, an error occurs. The number of columns inserted is equal to the number of columns selected.
- (void) insertFormulaFormula:(NSString *)formula numberFormat:(NSString *)numberFormat;  // Inserts an = Formula field that contains a formula at the selection. The formula replaces the selection, if the selection isn't collapsed.
- (void) insertRowsPosition:(MicrosoftWordEiAb)position numberOfRows:(NSInteger)numberOfRows;  // Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.
- (MicrosoftWordRevision *) nextRevisionWrap:(BOOL)wrap;  // Locates and returns the next tracked change as a revision object. The changed text becomes the current selection. Use the properties of the resulting revision object to see what type of change it is, who made it, and so forth.
- (void) pasteFormat;  // Applies formatting copied with the copy format method to the selection. If a paragraph mark was selected when the copy format method was used, Word applies paragraph formatting in addition to character formatting.
- (void) pasteObject;
- (MicrosoftWordRevision *) previousRevisionWrap:(BOOL)wrap;  // Locates and returns the previous tracked change as a revision object. The changed text becomes the current selection. Use the properties of the resulting revision object to see what type of change it is, who made it, and so forth.
- (void) selectCell;  // Selects the entire cell containing the current selection. To use this method, the current selection must be contained within a single cell.
- (void) selectColumn;  // Selects the column that contains the insertion point, or selects all columns that contain the selection. If the selection isn't in a table, an error occurs.
- (void) selectCurrentAlignment;  // Extends the selection forward until text with a different paragraph alignment is encountered.
- (void) selectCurrentColor;  // Extends the selection forward until text with a different color is encountered.
- (void) selectCurrentFont;  // Extends the selection forward until text in a different font or font size is encountered.
- (void) selectCurrentIndent;  // Extends the selection forward until text with different left or right paragraph indents is encountered.
- (void) selectCurrentSpacing;  // Extends the selection forward until a paragraph with different line spacing is encountered.
- (void) selectCurrentTabs;  // Extends the selection forward until a paragraph with different tab stops is encountered.
- (void) selectRow;  // Selects the row that contains the insertion point, or selects all rows that contain the selection. If the selection isn't in a table, an error occurs.
- (void) shrinkDiscontiguousSelection;  // Deselects all but the most recently selected text when a selection contains multiple, unconnected selections.
- (void) shrinkSelection;  // Shrinks the selection to the next smaller unit of text. The progression is as follows: entire document, section, paragraph, sentence, word, insertion point.
- (void) speakText;  // Have the selected text be spoken.
- (void) splitTableInSelection;  // Inserts an empty paragraph above the first row in the selection. If the selection isn't in the first row of the table, the table is split into two tables.
- (void) typeBackspace;  // Deletes the character preceding a collapsed selection, an insertion point. If the selection isn't collapsed to an insertion point, the selection is deleted.
- (void) typeParagraph;  // Inserts a new, blank paragraph. If the selection isn't collapsed to an insertion point, it's replaced by the new paragraph.
- (void) typeTextText:(NSString *)text;  // Inserts the specified text. If the Word options property replace selection is true, the selection is replaced by the specified text. If replace selection is false, the specified text is inserted before the selection.

@end

// Represents a subdocument within a document or range.
@interface MicrosoftWordSubdocument : MicrosoftWordBaseObject

@property (readonly) BOOL hasFile;  // Returns true if the specified subdocument has been saved to a file.
@property (readonly) NSInteger level;  // Returns the heading level used to create the subdocument.
@property BOOL locked;  // Returns or sets if a subdocument in a master document is locked.
@property (copy, readonly) NSString *name;  // The name of the subdocument.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified subdocument.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the subdocument.

- (MicrosoftWordDocument *) openSubdocument;  // Opens the specified subdocument. Returns a document object representing the opened object.
- (void) splitSubdocumentTextRange:(MicrosoftWordTextRange *)textRange;  // Divides an existing subdocument into two subdocuments at the same level in master document view or outline view. The division is at the beginning of the specified range.

@end

// Contains information about the computer system.
@interface MicrosoftWordSystemObject : MicrosoftWordBaseObject

@property (readonly) MicrosoftWordE118 country;  // Returns the country or region designation of the system
@property MicrosoftWordE139 cursor;  // Returns or sets the state of the pointer.
@property (readonly) NSInteger horizontalResolution;  // Returns the horizontal display resolution, in pixels.
@property (copy, readonly) NSString *operatingSystem;  // Returns the name of the current operating system.
@property (copy, readonly) NSString *processorType;  // Returns the type of processor that the system is using.
@property (copy, readonly) NSString *systemVersion;  // Returns the version number of the operating system.
@property (readonly) NSInteger verticalResolution;  // Returns the vertical display resolution, in pixels.

- (NSString *) getPrivateProfileStringFileName:(NSString *)fileName section:(NSString *)section key:(NSString *)key;  // Returns a string in a settings file.
- (NSString *) getProfileStringSection:(NSString *)section key:(NSString *)key;  // Returns a string from the application's settings file.
- (void) setPrivateProfileStringFileName:(NSString *)fileName section:(NSString *)section key:(NSString *)key privateProfileString:(NSString *)privateProfileString;  // Sets a string in a settings file.
- (void) setProfileStringSection:(NSString *)section key:(NSString *)key profileString:(NSString *)profileString;  // Sets a string from the application's settings file.

@end

// Represents a single tab stop. 
@interface MicrosoftWordTabStop : MicrosoftWordBaseObject

@property MicrosoftWordE145 alignment;  // Returns or sets a constant that represents the alignment for the specified tab stop.
@property (readonly) BOOL customTab;  // Returns if the specified tab stop is a custom tab stop.
@property (copy, readonly) MicrosoftWordTabStop *nextTabStop;  // Returns the next tab stop object
@property (copy, readonly) MicrosoftWordTabStop *previousTabStop;  // Returns the previous tab stop object
@property MicrosoftWordE170 tabLeader;  // Returns or sets the leader for the specified tab stop object
@property double tabStopPosition;  // Returns or sets the position of a tab stop relative to the left margin.


@end

// Represents a single table of authorities in a document
@interface MicrosoftWordTableOfAuthorities : MicrosoftWordBaseObject

@property NSInteger category;  // Returns or sets the category of entries to be included in a table of authorities.
@property (copy) NSString *entrySeparator;  // Returns or sets the characters up to five that separate a table of authorities entry and its page number. The default is a tab character with a dotted leader.
@property BOOL includeCategoryHeader;  // Returns or sets if the category name for a group of entries appears in the table of authorities.
@property (copy) NSString *includeSequenceName;  // Returns or sets the Sequence SEQ field identifier for a table of authorities.
@property BOOL keepEntryFormatting;  // Returns or sets if formatting from table of authorities entries is applied to the entries in the specified table of authorities.
@property (copy) NSString *pageNumberSeparator;  // Returns of sets the characters up to five that separate individual page references in a table of authorities. The default is a comma and a space.
@property (copy) NSString *pageRangeSeparator;  // Returns or sets the characters up to five that separate a range of pages in a table of authorities. The default is an en dash.
@property BOOL passim;  // Returns or sets if five or more page references to the same authority are replaced with Passim.
@property (copy) NSString *separator;  // Returns or sets the characters up to five between the sequence number and the page number. A hyphen is the default character.
@property MicrosoftWordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in a table of authorities.
@property (copy) NSString *tableOfAuthoritiesBookmark;  // Returns or sets the name of the bookmark from which to collect table of authorities entries.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the table of authorities object.


@end

// Represents a single table of contents in a document.
@interface MicrosoftWordTableOfContents : MicrosoftWordBaseObject

- (SBElementArray *) headingStyles;

@property BOOL includePageNumbers;  // Returns or sets if page numbers are included in the table of contents.
@property NSInteger lowerHeadingLevel;  // Returns or sets the ending heading level for a table of contents.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in a table of contents.
@property MicrosoftWordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in a table of contents.
@property (copy) NSString *tableId;  // Returns or sets a one-letter identifier that's used to build a table of contents or table of figures from TOC fields.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the table of contents object.
@property NSInteger upperHeadingLevel;  // Returns or sets the starting heading level for a table of contents.
@property BOOL useFields;  // Returns or sets if table of contents entry fields are used to create a table of contents.
@property BOOL useHeadingStyles;  // Returns or sets if built-in heading styles are used to create a table of contents.


@end

// Represents a single table of figures in a document.
@interface MicrosoftWordTableOfFigures : MicrosoftWordBaseObject

- (SBElementArray *) headingStyles;

@property (copy) NSString *caption;  // Returns or sets the label that identifies the items to be included in a table of figures.
@property BOOL includeLabel;  // Returns or sets if the caption label and caption number are included in a table of figures.
@property BOOL includePageNumbers;  // Returns or sets if page numbers are included in the table of figures.
@property NSInteger lowerHeadingLevel;  // Returns or sets the ending heading level for a table of figures.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in a table of figures.
@property MicrosoftWordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in a table of figures.
@property (copy) NSString *tableId;  // Returns or sets a one-letter identifier that's used to build a table of figures from TOC fields.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the table of figures object.
@property NSInteger upperHeadingLevel;  // Returns or sets the starting heading level for a table of figures.
@property BOOL useFields;  // Returns or sets if table of figures entry fields are used to create a table of figures.
@property BOOL useHeadingStyles;  // Returns or sets if built-in heading styles are used to create a table of figures.


@end

// Represents a document template.
@interface MicrosoftWordTemplate : MicrosoftWordBaseObject

- (SBElementArray *) autoTextEntries;
- (SBElementArray *) documentProperties;
- (SBElementArray *) customDocumentProperties;
- (SBElementArray *) listTemplates;
- (SBElementArray *) buildingBlocks;
- (SBElementArray *) buildingBlockTypes;

@property MicrosoftWordE108 eastAsianLineBreak;  // Returns or sets the line break control level for the specified template. This property is ignored if the “east asian line break control” property is set to false. Note that “east asian line break control” is a paragraph format property.
@property (copy, readonly) NSString *fullName;  // Specifies the name of the template including the drive or Web path.
@property MicrosoftWordE107 justificationMode;  // Returns or sets the character spacing adjustment for the specified template.
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the template object
@property MicrosoftWordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (copy, readonly) NSString *name;  // Returns the name of the template.
@property (copy) NSString *noLineBreakAfter;  // Returns or sets the kinsoku characters after which Microsoft Word will not break a line.
@property (copy) NSString *noLineBreakBefore;  // Returns or sets the kinsoku characters before which Microsoft Word will not break a line.
@property BOOL noProofing;  // Returns or sets if the spelling and grammar checker should ignore documents based on this template.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the template file.
@property BOOL saved;  // Return or set if the template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed.
@property (readonly) MicrosoftWordE101 templateType;  // Returns the template type.

- (MicrosoftWordBuildingBlock *) addBuildingBlockToTemplateName:(NSString *)name type:(MicrosoftWordE329)type category:(NSString *)category fromLocation:(MicrosoftWordTextRange *)fromLocation objectDescription:(NSString *)objectDescription insertOptions:(MicrosoftWordE330)insertOptions;  // Creates a new building block entry in a template and returns a building block object that represents the new building block entry.
- (MicrosoftWordAutoTextEntry *) appendToSpikeRange:(MicrosoftWordTextRange *)range;  // Deletes the specified range and adds the contents of the range to the Spike which is a built-in autotext entry.
- (MicrosoftWordDocument *) openAsDocument;  // Opens the specified template as a document and returns a document object. Opening a template as a document allows the user to edit the contents of the template.

@end

// Represents a single text column.
@interface MicrosoftWordTextColumn : MicrosoftWordBaseObject

@property double spaceAfter;  // Returns or sets the amount of spacing in points after the text column.
@property double width;  // Returns or sets the width of the object.


@end

// Represents a single text form field.
@interface MicrosoftWordTextInput : MicrosoftWordBaseObject

@property (copy) NSString *defaultTextInput;  // Returns or sets the text that represents the default text box contents.
@property (copy, readonly) NSString *format;  // Returns the formatting string for this text input object.
@property (readonly) MicrosoftWordE188 textInputFieldType;  // Returns the type of text form field.
@property (readonly) BOOL valid;  // Returns if the text input object is valid.
@property NSInteger width;  // Returns or sets the width of the object.

- (void) editTypeFormFieldType:(MicrosoftWordE188)formFieldType defaultType:(NSString *)defaultType typeFormat:(NSString *)typeFormat enabled:(BOOL)enabled;  // Sets options for the specified text form field.

@end

// Represents options that control how text is retrieved from a text range object.
@interface MicrosoftWordTextRetrievalMode : MicrosoftWordBaseObject

@property BOOL includeFieldCodes;  // Returns or sets if the text retrieved from the specified range includes field codes.
@property BOOL includeHiddenText;  // Returns or sets if the text retrieved from the specified range includes hidden text.
@property MicrosoftWordE202 viewType;  // Returns or sets the view for the text retrieval mode object.


@end

// Represents a variable stored as part of a document. Document variables are used to preserve macro settings in between macro sessions.
@interface MicrosoftWordVariable : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the variable
@property (copy) NSString *variableValue;  // Returns or sets the value of a variable as a string.


@end

// Contains the view attributes, show all, field shading, table gridlines, and so on, for a window or pane.
@interface MicrosoftWordView : MicrosoftWordBaseObject

- (SBElementArray *) reviewers;

@property BOOL browseToWindow;  // Returns or sets if lines wrap at the right edge of the window rather than at the right margin of the page.
@property BOOL conflictMode;  // Returns or sets the Conflict Mode.
@property BOOL dataMergeDataView;  // Returns or sets if data merge data is displayed instead of data merge fields in the specified window.
@property BOOL draft;  // Returns or sets if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.
@property NSInteger enlargeFontsLessThan;  // Returns or sets the point size below which screen fonts are automatically scaled to the larger size.
@property MicrosoftWordE229 fieldShading;  // Returns or sets on-screen shading for form fields.
@property BOOL fullScreen;  // Returns or sets if the window is in full-screen view.
@property BOOL magnifier;  // Returns or sets if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages.
@property MicrosoftWordE301 revisionsMode;  // Returns or sets the way Microsoft Word shows revisions.
@property MicrosoftWordE300 revisionsView;  // Returns or sets whether Microsoft Word shows revisions based on the final or the original document.
@property MicrosoftWordE203 seekView;  // Returns or sets the document element displayed in print layout view.
@property BOOL showAll;  // Returns or sets if all nonprinting characters such as hidden text, tab marks, space marks, and paragraph marks are displayed. 
@property BOOL showBookmarks;  // Returns or sets if square brackets are displayed at the beginning and end of each bookmark.
@property BOOL showComments;  // Returns or sets if Microsoft Word displays comments.
@property BOOL showDrawings;  // Returns or sets if objects created with the drawing tools are displayed in print layout view.
@property BOOL showFieldCodes;  // Returns or sets if field codes are displayed.
@property BOOL showFirstLineOnly;  // Returns or sets if only the first line of body text is shown in outline view.
@property BOOL showFormat;  // Returns or set if character formatting is visible in outline view.
@property BOOL showFormatChanges;  // Returns or sets if Microsoft Word displays formatting changes.
@property BOOL showHiddenText;  // Returns or set if text formatted as hidden text is displayed. 
@property BOOL showHighlight;  // Returns or sets if highlight formatting is displayed and printed with a document.
@property BOOL showHyphens;  // Returns or sets if optional hyphens are displayed. An optional hyphen indicates where to break a word when it falls at the end of a line.
@property BOOL showInsertionsAndDeletions;  // Returns or sets if Microsoft Word displays insertions and deletions.
@property BOOL showMainTextLayer;  // Returns or sets if the text in the specified document is visible when the header and footer areas are displayed.
@property BOOL showObjectAnchors;  // Returns or set if object anchors are displayed next to items that can be positioned in print layout view.
@property BOOL showOptionalBreaks;  // Returns or sets if Microsoft Word displays optional line breaks.
@property BOOL showOtherAuthors;  // Returns or sets whether Microsoft Word shows other authors who are also editing the document.
@property BOOL showParagraphs;  // Returns or sets if paragraph marks are displayed. 
@property BOOL showPicturePlaceHolders;  // Returns or sets if blank boxes are displayed as placeholders for pictures.
@property BOOL showRevisionsAndComments;  // Returns or sets if Microsoft Word displays revisions and comments.
@property BOOL showSpaces;  // Returns or sets if space characters are displayed.
@property BOOL showTabs;  // Returns or sets if tab characters are displayed.
@property BOOL showTextBoundaries;  // Returns or sets if dotted lines are displayed around page margins, text columns, objects, and frames in print layout view. 
@property MicrosoftWordE204 splitSpecial;  // Returns or sets the active window pane.
@property BOOL tableGridlines;  // Returns or sets if table gridlines are displayed. 
@property MicrosoftWordE202 viewType;  // Returns or sets the view type.
@property BOOL wrapToWindow;  // Returns or sets if lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.
@property (copy, readonly) MicrosoftWordZoom *zoom;  // Returns the zoom object associated with this view object.

- (void) collapseOutlineTextRange:(MicrosoftWordTextRange *)textRange;  // Collapses the text under the selection or the specified range by one heading level.
- (void) expandOutlineTextRange:(MicrosoftWordTextRange *)textRange;  // Expands the text under the selection or the specified range by one heading level.
- (void) nextHeaderFooter;  // If the selection is in a header, this method moves to the next header within the current section or to the first header in the following section. If the selection is in a footer, this method moves to the next footer.
- (void) previousHeaderFooter;  // If the selection is in a header, this method moves to the previous header within the current section or to the last header in the previous section. If the selection is in a footer, this method moves to the previous footer.
- (void) showAllHeadings;  // Toggles between showing all text, headings and body text, and showing only headings.
- (void) showHeadingLevel:(NSInteger)level;  // Shows all headings up to the specified heading level and hides subordinate headings and body text. This method generates an error if the view isn't outline view or master document view.

@end

// Contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page.
@interface MicrosoftWordWebOptions : MicrosoftWordBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page.
@property (copy, readonly) NSString *docKeywords;  // Returns the keywords associated with a document.
@property (copy, readonly) NSString *docTitle;  // Returns the title for a Web document.
@property MicrosoftWordMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property BOOL roundTripHTML;  // Returns or sets if whether to save an HTML document with information that is specific to Microsoft Word. Setting this property to true allows you to preserve all Word settings in an HTML document.
@property MicrosoftWordMSsz screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL useLongFileNames;  // Returns or sets if long file names are used when you save the document as a Web page.

- (void) useDefaultFolderSuffix;  // Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

@end

// Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.
@interface MicrosoftWordWindow : MicrosoftWordBasicWindow

- (SBElementArray *) panes;

@property MicrosoftWordE103 IMEMode;  // Returns or sets the default start-up mode for the Japanese Input Method Editor, IME
@property (readonly) BOOL active;  // Returns if the window is active.
@property (copy, readonly) MicrosoftWordPane *activePane;  // Returns a pane object that represents the active pane for the window 
@property (copy) NSString *caption;  // Returns or sets the caption text for the document window.
@property BOOL displayHorizontalScrollBar;  // Returns or sets if the horizontal scroll bar is visible
@property BOOL displayRulers;  // Returns or sets if rulers are displayed for the window
@property BOOL displayScreenTips;  // Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted.
@property BOOL displayVerticalRuler;  // Returns or sets if vertical rulers are displayed for the window
@property BOOL displayVerticalScrollBar;  // Returns or sets if the vertical scroll bar is visible
@property (copy, readonly) MicrosoftWordDocument *document;  // Returns a document object associated with the pane. 
@property BOOL documentMap;  // Returns or sets if the document map is visible.
@property NSInteger documentMapPercentWidth;  // Returns or sets the width of the document map as a percentage of the width of the specified window.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property NSInteger height;  // Returns or sets the height of the object.
@property NSInteger horizontalPercentScrolled;  // Returns or sets the horizontal scroll position as a percentage of the document width.
@property NSInteger leftPosition;  // Returns or sets the horizontal position of the object.
@property (copy, readonly) MicrosoftWordWindow *nextWindow;  // Returns the next window object
@property (copy, readonly) MicrosoftWordWindow *previousWindow;  // Returns the previous window object
@property (copy, readonly) MicrosoftWordSelectionObject *selection;  // Returns the selection object that represents a selected text range or the insertion point.
@property NSInteger splitVertical;  // Returns or sets the vertical split percentage for the specified window. To remove the split, set this property to zero or set the split window property to false.
@property BOOL splitWindow;  // Returns or sets if the window is split into multiple panes. When setting this to true the window is split into two equal-sized window panes.
@property double styleAreaWidth;  // Returns or sets the width of the style area in points.
@property NSInteger top;  // Returns or sets the vertical position of the object.
@property NSInteger verticalPercentScrolled;  // Returns or sets the vertical scroll position as a percentage of the document width.
@property (copy, readonly) MicrosoftWordView *view;  // Returns a view object that represents the view for the window.
@property NSInteger width;  // Returns or sets the width of the object.
@property (readonly) NSInteger windowNumber;  // Returns the window number of the document displayed in the specified window. For example, if the caption of the window is Sales.doc:2, this property returns the number 2.
@property MicrosoftWordE198 windowState;  // Returns or sets the state of the window.
@property (readonly) MicrosoftWordE201 windowType;  // Returns the window type.

- (void) toggleShowAllReviewers;  // Toggles whether or not all reviewers are shown in this window.

@end

// Contains magnification options, for example, the zoom percentage for a window or pane.
@interface MicrosoftWordZoom : MicrosoftWordBaseObject

@property NSInteger pageColumns;  // Returns or sets the number of pages to be displayed side by side on-screen at the same time in print layout view or print preview.
@property MicrosoftWordE205 pageFit;  // Returns or sets the view magnification of a window so that either the entire page is visible or the entire width of the page is visible.
@property NSInteger pageRows;  // Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. 
@property NSInteger percentage;  // Returns or sets the magnification for a window as a percentage.


@end



/*
 * Drawing Suite
 */

@interface MicrosoftWordAdjustment : MicrosoftWordBaseObject

@property double adjustment_value;  // Returns or sets a floating point adjustment for a shape.


@end

// Contains properties and methods that apply to line callouts.
@interface MicrosoftWordCalloutFormat : MicrosoftWordBaseObject

@property BOOL accent;  // Returns or sets if a vertical accent bar separates the callout text from the callout line.
@property MicrosoftWordMCAt angle;  // Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box.
@property BOOL autoAttach;  // Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box.
@property (readonly) BOOL autoLength;  // Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false.
@property (readonly) double calloutFormatLength;  // When the AutoLength property of the specified callout is set to False, the Length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box.
@property BOOL calloutHasBorder;  // Returns or sets whether the text in the specified callout is surrounded by a border.
@property MicrosoftWordMCot calloutType;  // Returns or sets the callout type.
@property (readonly) double drop;  // For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
@property (readonly) MicrosoftWordMCDt dropType;  // Returns a value that indicates where the callout line attaches to the callout text box.
@property double gap;  // Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box.


@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface MicrosoftWordFillFormat : MicrosoftWordBaseObject

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property MicrosoftWordMCoI backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) MicrosoftWordMFdT fillType;  // Returns the shape fill format type.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property MicrosoftWordMCoI foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) MicrosoftWordMGCt gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend.
@property (readonly) MicrosoftWordMGdS gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2.
@property (readonly) MicrosoftWordPpTy pattern;  // Returns the value that represents the pattern applied to the specified fill or line.
@property (readonly) MicrosoftWordMPGb presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) MicrosoftWordMPzT presetTexture;  // Returns the preset texture for the specified fill.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property (readonly) MicrosoftWordMxtT textureType;  // Returns the texture type for the specified fill.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Represents the glow formatting for a shape or range of shapes.
@interface MicrosoftWordGlowFormat : MicrosoftWordBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified glow format.
@property MicrosoftWordMCoI colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents horizontal line formatting.
@interface MicrosoftWordHorizontalLineFormat : MicrosoftWordBaseObject

@property MicrosoftWordE280 alignment;  // Returns or sets a constant that represents the alignment for the specified horizontal line.
@property BOOL noShade;  // Returns or sets if Microsoft Word draws the specified horizontal line without 3-D shading.
@property double percentWidth;  // Returns or sets the length of the specified horizontal line expressed as a percentage of the window width.
@property MicrosoftWordE281 widthType;  // Returns or sets the width type for the specified horizontal line format object.


@end

// Represents a graphic object in the text layer of a document.
@interface MicrosoftWordInlineShape : MicrosoftWordBaseObject

@property (copy) NSString *alternativeText;  // Returns or sets the alternative text associated with a shape in a Web page.
@property (readonly) NSInteger anchorID;  // Returns the anchor id for the specified shape.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with the inline shape
@property (readonly) NSInteger editID;  // Returns the edit id for the specified shape.
@property (copy, readonly) MicrosoftWordField *field;  // Returns a field object that represents the field associated with the specified inline shape.
@property (copy, readonly) MicrosoftWordFillFormat *fillFormat;  // Returns the fill format object associated with this inline shape object.
@property (copy, readonly) MicrosoftWordGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property double height;  // Returns or sets the height of the object.
@property (copy, readonly) MicrosoftWordHorizontalLineFormat *horizontalLineFormat;  // Returns the horizontal line format object associated with this inline shape object.
@property (copy, readonly) MicrosoftWordHyperlinkObject *hyperlink;  // Returns the hyperlink object associated with this inline shape object.
@property double inlineShapeScaleHeight;  // Returns or sets the height of the specified inline shape relative to its original size. 
@property double inlineShapeScaleWidth;  // Returns or sets the width of the specified inline shape relative to its original size. 
@property (readonly) MicrosoftWordE259 inlineShapeType;  // The type of this inline shape.
@property BOOL isInlinePlaceholder;  // Returns true if a shape is a placeholder.
@property (readonly) BOOL isPictureBullet;  // Returns true if an InlineShape object is a picture bullet.
@property (copy, readonly) MicrosoftWordLineFormat *lineFormat;  // Returns the line format object associated with this inline shape object.
@property (copy, readonly) MicrosoftWordLinkFormat *linkFormat;  // Returns the link format object associated with this inline shape object.
@property BOOL lockAspectRatio;  // Returns or set if the specified shape retains its original proportions when you resize it.
@property (copy, readonly) MicrosoftWordPictureFormat *pictureFormat;  // Returns the picture format object associated with this inline shape object.
@property (copy, readonly) MicrosoftWordReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property (copy, readonly) MicrosoftWordShadowFormat *shadow;  // Returns the shadow format object associated with this shape object.
@property (copy, readonly) MicrosoftWordSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the text for the inline shape.
@property double width;  // Returns or sets the width of the object.
@property (copy, readonly) MicrosoftWordWordArtFormat *wordArtFormat;  // Returns the word art format object associated with the word art shape object.

- (MicrosoftWordShape *) convertToShape;  // Converts an inline shape to a free-floating shape.
- (MicrosoftWordBorder *) getBorderWhichBorder:(MicrosoftWordE122)whichBorder;  // Returns the specified border object.
- (void) reset;  // Removes changes that were made to an inline shape.

@end

// Represents a horizontal line object in the text layer of a document.
@interface MicrosoftWordInlineHorizontalLine : MicrosoftWordInlineShape

@property (copy, readonly) NSString *fileName;  // The file name that contains the picture of the horizontal line.


@end

// Represents a picture bullet object in the text layer of a document.
@interface MicrosoftWordInlinePictureBullet : MicrosoftWordInlineShape

@property (copy, readonly) NSString *fileName;  // The file name that contains the picture of the picture bullet.


@end

// Represents a picture object in the text layer of a document.
@interface MicrosoftWordInlinePicture : MicrosoftWordInlineShape

@property (copy, readonly) NSString *fileName;  // The file name that contains the picture.
@property (readonly) BOOL linkToFile;  // Returns true if the picture shape is linked to the file.
@property (copy, readonly) MicrosoftWordPictureFormat *pictureFormat;  // Returns the picture format object associated with this inline picture object.
@property (readonly) BOOL saveWithDocument;  // Specifies if the picture should be saved with the document.


@end

// Represents line and arrowhead formatting. For a line, the line format object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.
@interface MicrosoftWordLineFormat : MicrosoftWordBaseObject

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property MicrosoftWordMCoI backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property MicrosoftWordMAhL beginArrowheadLength;  // Returns or sets the length of the arrowhead at the beginning of the specified line.
@property MicrosoftWordMAhS beginArrowheadStyle;  // Returns or sets the style of the arrowhead at the beginning of the specified line.
@property MicrosoftWordMAhW beginArrowheadWidth;  // Returns or sets the width of the arrowhead at the beginning of the specified line.
@property MicrosoftWordMlDs dashStyle;  // Returns or sets the dash style for the specified line.
@property MicrosoftWordMAhL endArrowheadLength;  // Returns or sets the length of the arrowhead at the end of the specified line.
@property MicrosoftWordMAhS endArrowheadStyle;  // Returns or sets the style of the arrowhead at the end of the specified line.
@property MicrosoftWordMAhW endArrowheadWidth;  // Returns or sets the width of the arrowhead at the end of the specified line.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property MicrosoftWordMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property MicrosoftWordMLnS lineStyle;  // Returns or sets the line format style.
@property MicrosoftWordPpTy pattern;  // Returns or sets a value that represents the pattern applied to the specified fill or line.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.
@property double weight;  // Returns or sets the thickness of the specified line in points.


@end

// Represents a Microsoft Office theme.
@interface MicrosoftWordOfficeTheme : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) MicrosoftWordThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) MicrosoftWordThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

// Contains properties and methods that apply to pictures. 
@interface MicrosoftWordPictureFormat : MicrosoftWordBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSColor *transparencyColor;  // Returns or sets the transparent color for the specified picture as a RGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

// Represents the reflection effect in Office graphics.
@interface MicrosoftWordReflectionFormat : MicrosoftWordBaseObject

@property MicrosoftWordMRfT reflectionType;  // Returns or sets the type of the reflection format object.


@end

// Represents shadow formatting for a shape.
@interface MicrosoftWordShadowFormat : MicrosoftWordBaseObject

@property double blur;  // Returns or sets the blur, in points, of the specified shadow.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property MicrosoftWordMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;  // Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
@property double offsetX;  // Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
@property double offsetY;  // Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape.
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property MicrosoftWordMSSt shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property MicrosoftWordMSdT shadowType;  // Returns or sets the shape shadow type.
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Represents an object in the drawing layer.
@interface MicrosoftWordShape : MicrosoftWordBaseObject

- (SBElementArray *) adjustments;

@property (copy, readonly) MicrosoftWordTextRange *anchor;  // Returns a text range object that represents the anchoring range for the specified shape.
@property (readonly) NSInteger anchorID;  // Returns the anchor id for the specified shape.
@property MicrosoftWordMAsT autoShapeType;  // Returns or sets the shape type for the specified shape object, which must represent an autoshape.
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger editID;  // Returns the edit id for the specified shape.
@property (copy, readonly) MicrosoftWordFillFormat *fillFormat;  // Return the fill format object associated with this shape object.
@property (copy, readonly) MicrosoftWordGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property BOOL hasChart;  // True if the specified shape has a chart.
@property double height;  // Returns or sets the height of the object.
@property (readonly) BOOL horizontalFlip;  // Returns true if the shape has been flipped horizontally. 
@property (copy, readonly) MicrosoftWordHyperlinkObject *hyperlink;  // Returns the hyperlink object associated with this shape object.
@property BOOL isPlaceholder;  // Returns true if a shape is a placeholder.
@property double leftPosition;  // Returns or sets the horizontal position of the object.
@property (copy, readonly) MicrosoftWordLineFormat *lineFormat;  // Returns the line format object associated with this shape object.
@property (copy, readonly) MicrosoftWordLinkFormat *linkFormat;  // Returns the link format object associated with this shape object.
@property BOOL lockAnchor;  // Returns or sets if the specified shape object's anchor is locked to the anchoring range. When a shape has a locked anchor, you cannot move the shape's anchor by dragging it, the anchor doesn't move as the shape is moved.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it.
@property (copy) NSString *name;  // Returns or sets the name of this shape object.
@property (copy, readonly) MicrosoftWordReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property MicrosoftWordE236 relativeHorizontalPosition;  // Returns or sets if the horizontal position of the shape is relative.
@property MicrosoftWordE237 relativeVerticalPosition;  // Returns or sets if the vertical position of the shape is relative.
@property double rotation;  // Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation.
@property (copy, readonly) MicrosoftWordShadowFormat *shadow;  // Returns the shadow format object associated with this shape object.
@property (readonly) MicrosoftWordMShp shapeType;  // Returns the shape type.
@property (copy, readonly) MicrosoftWordSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) MicrosoftWordTextFrame *textFrame;  // Returns the text frame object associated with this shape object.
@property (copy, readonly) MicrosoftWordThreeDFormat *threeDFormat;  // Returns the threeD format object associated with this shape object.
@property double top;  // Returns or sets the vertical position of the object.
@property (readonly) BOOL verticalFlip;  // Returns true if the shape has been flipped vertically.
@property BOOL visible;  // Returns or sets if the shape object is visible.
@property double width;  // Returns or sets the width of the object.
@property (copy, readonly) MicrosoftWordWrapFormat *wrapFormat;  // Returns the wrap format object associated with this shape object.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order.

- (void) apply;  // Applies to the specified shape formatting that has been copied using the pick up method.
- (MicrosoftWordFrame *) convertToFrame;  // Converts the specified shape to a frame. Returns a frame object that represents the new frame.
- (MicrosoftWordInlineShape *) convertToInlineShape;  // Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only picture shapes.
- (void) flipFlipCommand:(MicrosoftWordMFlC)flipCommand;  // Flips a shape horizontally or vertically.
- (void) pickUp;  // Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.
- (void) rerouteConnections;  // Reroutes the connections between shapes.
- (void) setShapesDefaultProperties;  // Applies the formatting of the specified shape to a default shape for that document. New shapes inherit many of their attributes from the default shape.
- (void) zOrderZOrderCommand:(MicrosoftWordMZoC)zOrderCommand;  // Moves the specified shape in front of or behind other shapes.

@end

// Represents an callout object in the drawing layer.
@interface MicrosoftWordCallout : MicrosoftWordShape

@property (copy, readonly) MicrosoftWordCalloutFormat *calloutFormat;  // Returns the callout format object associated with this callout shape object.
@property (readonly) MicrosoftWordMCot calloutType;  // Return the type of this callout


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface MicrosoftWordLineShape : MicrosoftWordShape

@property double beginLineX;  // Returns or sets the starting X coordinate for the line shape.
@property double beginLineY;  // Returns or sets the starting Y coordinate for the line shape.
@property double endLineX;  // Returns or sets the ending X coordinate for the line shape.
@property double endLineY;  // Returns or sets the ending Y coordinate for the line shape.


@end

// Represents an picture object in the drawing layer.
@interface MicrosoftWordPicture : MicrosoftWordShape

@property (copy, readonly) NSString *fileName;  // The name of the file containing the picture.
@property (readonly) BOOL linkToFile;  // Returns true if the picture shape is linked to the file.
@property (copy, readonly) MicrosoftWordPictureFormat *pictureFormat;  // Returns the picture format object associated with this picture shape.
@property (readonly) BOOL saveWithDocument;  // Specifies if the picture should be saved with the document.

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(MicrosoftWordMSFr)scale;  // Scales the height of the picture shape by a specified factor.
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(MicrosoftWordMSFr)scale;  // Scales the width of the shape by a specified factor.

@end

// Represents the soft edge formatting for a shape or range of shapes.
@interface MicrosoftWordSoftEdgeFormat : MicrosoftWordBaseObject

@property MicrosoftWordMSET softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Represents a standard horizontal line object in the text layer of a document.
@interface MicrosoftWordStandardInlineHorizontalLine : MicrosoftWordInlineShape


@end

// Represents an text box object in the drawing layer.
@interface MicrosoftWordTextBox : MicrosoftWordShape

@property (readonly) MicrosoftWordTxOr textOrientation;  // Returns the orientation of the text inside the text shape.


@end

// Represents the text frame in a shape object. Contains the text in the text frame as well as the properties that control the margins and orientation of the text frame.
@interface MicrosoftWordTextFrame : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordTextRange *containingRange;  // Returns a text range object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to.
@property (readonly) BOOL hasText;  // Returns true if the specified shape has text associated with it.
@property double marginBottom;  // Returns or sets the distance in points between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance in points between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance in points between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance in points between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property (copy) MicrosoftWordTextFrame *nextTextframe;  // Returns or sets the next text frame object.
@property MicrosoftWordTxOr orientation;  // Returns or sets the orientation of the text inside the frame.
@property (readonly) BOOL overflowing;  // Returns if the text inside the specified text frame doesn't all fit within the frame.
@property (copy) MicrosoftWordTextFrame *previousTextframe;  // Returns or sets the previous text frame object.
@property (copy, readonly) MicrosoftWordTextRange *textRange;  // Returns a text range object that represents the text range for this text frame.
@property BOOL textwrappingAllowed;  // Allow textwrapping of overlay objects.

- (void) breakForwardLink;  // Breaks the forward link for the specified text frame, if such a link exists.
- (BOOL) validLinkTargetTargetTextframe:(MicrosoftWordTextFrame *)targetTextframe;  // Determines whether the text frame of one shape can be linked to the text frame of another shape. Returns false if target textframe already contains text or is already linked, or if the shape doesn't support attached text.

@end

// Represents the color scheme of a Microsoft Office 2007 theme.
@interface MicrosoftWordThemeColorScheme : MicrosoftWordBaseObject

- (SBElementArray *) themeColors;

- (NSColor *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface MicrosoftWordThemeColor : MicrosoftWordBaseObject

@property (copy) NSColor *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) MicrosoftWordMCSI themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface MicrosoftWordThemeEffectScheme : MicrosoftWordBaseObject

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file

@end

// Represents the font scheme of a Microsoft Office theme.
@interface MicrosoftWordThemeFontScheme : MicrosoftWordBaseObject

- (SBElementArray *) minorThemeFonts;
- (SBElementArray *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface MicrosoftWordThemeFont : MicrosoftWordBaseObject

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface MicrosoftWordMajorThemeFont : MicrosoftWordThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface MicrosoftWordMinorThemeFont : MicrosoftWordThemeFont


@end

// Represents a collection of major and minor fonts in the font scheme of a Microsoft Office 2007 theme.
@interface MicrosoftWordThemeFonts : MicrosoftWordBaseObject


@end

// Represents a shape's three-dimensional formatting.
@interface MicrosoftWordThreeDFormat : MicrosoftWordBaseObject

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property MicrosoftWordMBlT bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property MicrosoftWordMBlT bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy) NSColor *contourColor;  // Returns or sets the color of the contour of an object or text.
@property MicrosoftWordMCoI contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy) NSColor *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property MicrosoftWordMCoI extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property MicrosoftWordMExC extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property MicrosoftWordMPzC presetCamera;  // Returns a constant that represents the camera preset.
@property MicrosoftWordMExD presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property MicrosoftWordMPLd presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property MicrosoftWordMLtT presetLightingRig;  // Returns a constant that represents the lighting preset.
@property MicrosoftWordMlSf presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property MicrosoftWordMPMt presetMaterial;  // Returns or sets the extrusion surface material.
@property MicrosoftWordM3DF presetThreeDFormat;  // Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Contains properties and methods that apply to WordArt objects.
@interface MicrosoftWordWordArtFormat : MicrosoftWordBaseObject

@property MicrosoftWordMTxA alignment;  // Returns or sets a constant that represents the alignment for the specified text effect.
@property BOOL bold;  // Returns or sets if the text of the word art shape is formatted as bold.
@property (copy) NSString *fontName;  // Returns or sets the font name of the font used by this word art shape.
@property double fontSize;  // Returns or sets the font size of the font used by this word art shape.
@property BOOL italic;  // Returns or sets if the text of the word art shape is formatted as italic.
@property BOOL kernedPairs;  // Returns or sets if character pairs in a WordArt object have been kerned. 
@property BOOL normalizedHeight;  // Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height.
@property MicrosoftWordMPTs presetShape;  // Returns or sets the shape of the specified WordArt.
@property MicrosoftWordMPXF presetWordArtEffect;  // Returns or sets the style of the specified WordArt.
@property BOOL rotatedChars;  // Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape.
@property double tracking;  // Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5.
@property (copy) NSString *wordArtText;  // Returns or sets the text associated with this word art shape.

- (void) toggleVerticalText;  // Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.

@end

// Represents an word art object in the drawing layer.
@interface MicrosoftWordWordArt : MicrosoftWordShape

@property (readonly) BOOL bold;  // Returns if the text of the word art shape is formatted as bold.
@property (copy, readonly) NSString *fontName;  // Returns the font name of the font used by this word art shape.
@property (readonly) double fontSize;  // Returns the font size of the font used by this word art shape.
@property (readonly) BOOL italic;  // Returns if the text of the word art shape is formatted as italic.
@property (readonly) MicrosoftWordMPXF presetWordArtEffect;  // Returns the style of the specified word art.
@property (copy, readonly) MicrosoftWordWordArtFormat *wordArtFormat;  // Returns the word art format object associated with the word art shape object.
@property (copy, readonly) NSString *wordArtText;  // The text associated with this word art shape.


@end

// Represents all the properties for wrapping text around a shape.
@interface MicrosoftWordWrapFormat : MicrosoftWordBaseObject

@property BOOL allowOverlap;  // Returns or sets a value that specifies whether a given shape can overlap other shapes.
@property double distanceBottom;  // Returns or sets the distance in points between the document text and the bottom edge of the text-free area surrounding the specified shape.
@property double distanceLeft;  // Returns or sets the distance in points between the document text and the left edge of the text-free area surrounding the specified shape.
@property double distanceRight;  // Returns or sets the distance in points between the document text and the right edge of the text-free area surrounding the specified shape.
@property double distanceTop;  // Returns or sets the distance in points between the document text and the top edge of the text-free area surrounding the specified shape.
@property MicrosoftWordE268 wrapSide;  // Returns or sets a value that indicates whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.
@property MicrosoftWordE267 wrapType;  // Returns or sets the wrap type for the specified shape.


@end



/*
 * Text Suite
 */

// Represents a single built-in or user-defined style. The Word style object includes style attributes, font, font style, paragraph spacing, and so on, as properties of the Word style object.
@interface MicrosoftWordWordStyle : MicrosoftWordBaseObject

@property BOOL automaticallyUpdate;  // True if the style is automatically redefined based on the selection. False if Word prompts for confirmation before redefining the style based on the selection.
@property MicrosoftWordE184 baseStyle;  // Returns or sets an existing style on which you can base the formatting of another style.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property (readonly) BOOL builtIn;  // Returns true if the specified object is one of the built-in styles in Word.
@property (copy, readonly) NSString *objectDescription;  // Returns the description of the specified style. For example, a typical description for the Heading 2 style might be Normal, Font: Arial, 12 pt, Bold, Italic, Space Before 12 pt After 3 pt, KeepWithNext, Level 2.
@property (copy, readonly) MicrosoftWordFont *fontObject;  // Returns the font object associated with this Word style object.
@property (copy, readonly) MicrosoftWordFrame *frame;  // Returns the frame object associated with this Word style object.
@property (readonly) BOOL inUse;  // Returns true if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the Word style object
@property MicrosoftWordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (readonly) NSInteger listLevelNumber;  // Returns the list level for the specified style.
@property (copy, readonly) MicrosoftWordListTemplate *listTemplate;  // Returns the list template object associated with this Word style object.
@property (copy) NSString *nameLocal;  // Returns or sets the name of a built-in style in the language of the user. Setting this property renames a user-defined style or adds an alias to a built-in style. 
@property MicrosoftWordE184 nextParagraphStyle;  // Returns or sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the specified style.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores for this Word style
@property BOOL noSpaceBetweenSame;  // Returns or sets if Microsoft Word suppresses space between paragraphs of the same style
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this Word style object.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the Word style.
@property (readonly) MicrosoftWordE128 styleType;  // Returns the style type.
@property (copy, readonly) MicrosoftWordTableStyle *tableStyle;  // Returns table style properties for this style.

- (void) linkToListTemplateListTemplate:(MicrosoftWordListTemplate *)listTemplate listLevelNumber:(NSInteger)listLevelNumber;  // Links the specified style to a list template so that the style's formatting can be applied to lists.

@end

// Represents all the formatting for a paragraph.
@interface MicrosoftWordParagraphFormat : MicrosoftWordBaseObject

- (SBElementArray *) tabStops;

@property BOOL addSpaceBetweenEastAsianAndAlpha;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs.
@property BOOL addSpaceBetweenEastAsianAndDigit;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs.
@property MicrosoftWordE142 alignment;  // Returns or sets a constant that represents the alignment for the specified paragraphs.
@property BOOL autoAdjustRightIndent;  // Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you’ve specified a set number of characters per line.
@property MicrosoftWordE104 baseLineAlignment;  // Returns or sets a constant that represents the vertical position of fonts on a line.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property double characterUnitFirstLineIndent;  // Returns or sets the value in characters for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property double characterUnitLeftIndent;  // Returns or sets the left indent value in characters for the specified paragraphs.
@property double characterUnitRightIndent;  // Returns or sets the right indent value in characters for the specified paragraphs.
@property BOOL disableLineHeightGrid;  // Returns or sets if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified.
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value in points for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property BOOL halfWidthPunctuationOnTopOfLine;  // Returns or sets if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs.
@property BOOL hangingPunctuation;  // Returns or sets if hanging punctuation is enabled for the specified paragraphs.
@property BOOL hyphenation;  // Returns or sets if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation.
@property BOOL keepTogether;  // Returns or sets if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
@property BOOL keepWithNext;  // Returns or sets if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
@property double lineSpacing;  // Returns or sets the line spacing in points for the specified paragraphs.
@property MicrosoftWordE157 lineSpacingRule;  // Returns or sets the line spacing for the specified paragraphs.
@property double lineUnitAfter;  // Returns or sets the amount of spacing in gridlines after the specified paragraphs.
@property double lineUnitBefore;  // Returns or sets the amount of spacing in gridlines before the specified paragraphs.
@property BOOL noLineNumber;  // Returns or set if line numbers are repressed for the specified paragraphs.
@property MicrosoftWordE269 outlineLevel;  // Returns or sets the outline level for the specified paragraphs.
@property BOOL pageBreakBefore;  // Returns or sets if a page break is forced before the specified paragraphs.
@property double paragraphFormatLeftIndent;  // Returns or sets the left indent in points for the specified paragraphs.
@property double paragraphFormatRightIndent;  // Returns or sets the right indent in points for the specified paragraphs.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the paragraph format object.
@property double spaceAfter;  // Returns or sets the spacing in points after the specified paragraphs.
@property BOOL spaceAfterAuto;  // Returns or sets if Microsoft Word automatically sets the amount of spacing after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing in points before the specified paragraphs.
@property BOOL spaceBeforeAuto;  // Returns or sets if Microsoft Word automatically sets the amount of spacing before the specified paragraphs.
@property MicrosoftWordE184 style;  // Returns or sets the Word style associated with the replacement object.
@property BOOL widowControl;  // Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document.
@property BOOL wordWrap;  // Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames.


@end

// Represents a single paragraph in a selection, range, or document.
@interface MicrosoftWordParagraph : MicrosoftWordBaseObject

- (SBElementArray *) tabStops;

@property BOOL addSpaceBetweenEastAsianAndAlpha;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs.
@property BOOL addSpaceBetweenEastAsianAndDigit;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs.
@property MicrosoftWordE142 alignment;  // Returns or sets a constant that represents the alignment for the specified paragraphs.
@property BOOL autoAdjustRightIndent;  // Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you’ve specified a set number of characters per line.
@property MicrosoftWordE104 baseLineAlignment;  // Returns or sets a constant that represents the vertical position of fonts on a line.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property double characterUnitFirstLineIndent;  // Returns or sets the value in characters for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property double characterUnitLeftIndent;  // Returns or sets the left indent value in characters for the specified paragraphs.
@property double characterUnitRightIndent;  // Returns or sets the right indent value in characters for the specified paragraphs.
@property BOOL disableLineHeightGrid;  // Returns or sets if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified.
@property (copy, readonly) MicrosoftWordDropCap *dropCap;  // Returns the drop cap object associated with this paragraph object.
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value in points for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property BOOL halfWidthPunctuationOnTopOfLine;  // Returns or sets if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs.
@property BOOL hangingPunctuation;  // Returns or sets if hanging punctuation is enabled for the specified paragraphs.
@property BOOL hyphenation;  // Returns or sets if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation.
@property BOOL keepTogether;  // Returns or sets if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
@property BOOL keepWithNext;  // Returns or sets if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
@property double lineSpacing;  // Returns or sets the line spacing in points for the specified paragraphs.
@property MicrosoftWordE157 lineSpacingRule;  // Returns or sets the line spacing for the specified paragraphs.
@property double lineUnitAfter;  // Returns or sets the amount of spacing in gridlines after the specified paragraphs.
@property double lineUnitBefore;  // Returns or sets the amount of spacing in gridlines before the specified paragraphs.
@property BOOL noLineNumber;  // Returns or set if line numbers are repressed for the specified paragraphs.
@property MicrosoftWordE269 outlineLevel;  // Returns or sets the outline level for the specified paragraphs.
@property BOOL pageBreakBefore;  // Returns or sets if a page break is forced before the specified paragraphs.
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this paragraph object.
@property double paragraphLeftIndent;  // Returns or sets the left indent in points for the specified paragraphs.
@property NSInteger paragraph_id;  // Returns the paragraph id for the specified paragraphs.
@property double rightIndent;  // Returns or sets the right indent in points for the specified paragraphs.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the paragraph object.
@property double spaceAfter;  // Returns or sets the spacing in points after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing in points before the specified paragraphs.
@property MicrosoftWordE184 style;  // Returns or sets the Word style associated with the replacement object.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the section object.
@property NSInteger text_id;  // Returns the text id for the specified paragraphs.
@property BOOL widowControl;  // Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document.
@property BOOL wordWrap;  // Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames.

- (void) indent;  // Indents one or more paragraphs by one level.
- (MicrosoftWordParagraph *) nextParagraph;  // Returns the next paragraph object.
- (void) outdent;  // Removes one level of indent for one or more paragraphs.
- (void) outlineDemote;  // Applies the next heading level style Heading 1 through Heading 8 to the specified paragraph or paragraphs. For example, if a paragraph is formatted with the Heading 2 style, this method demotes the paragraph by changing the style to Heading 3.
- (void) outlineDemoteToBody;  // Demotes the specified paragraph or paragraphs to body text by applying the Normal style.
- (void) outlinePromote;  // Applies the previous heading level style Heading 1 through Heading 8 to the specified paragraph or paragraphs. For example, if a paragraph is formatted with the Heading 2 style, this method promotes the paragraph by changing the style to Heading 1.
- (MicrosoftWordParagraph *) previousParagraph;  // Returns the previous paragraph object.

@end

// Represents a single section in a selection, range, or document.
@interface MicrosoftWordSection : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;
@property (copy) MicrosoftWordPageSetup *pageSetup;  // Returns or sets a page setup object associated with this section object
@property BOOL protectedForForms;  // Returns or sets if the specified section is protected for forms. When a section is protected for forms, you can select and modify text only in form fields.
@property (readonly) NSInteger sectionIndex;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the section object.

- (MicrosoftWordHeaderFooter *) getFooterIndex:(MicrosoftWordE163)index;  // Returns a specific footer object
- (MicrosoftWordHeaderFooter *) getHeaderIndex:(MicrosoftWordE163)index;  // Returns a specific header object

@end

// Contains shading attributes for an object.
@interface MicrosoftWordShading : MicrosoftWordBaseObject

@property (copy) NSColor *backgroundPatternColor;  // Returns or sets the RGB color that's applied to the background of the shading object.
@property MicrosoftWordE110 backgroundPatternColorIndex;  // Returns or sets the color index that's applied to the background of the shading object.
@property MicrosoftWordMCoI backgroundPatternColorThemeIndex;  // Returns or sets the color that's applied to the background of the shading object.
@property (copy) NSColor *foregroundPatternColor;  // Returns or sets the RGB color that's applied to the foreground of the shading object.
@property MicrosoftWordE110 foregroundPatternColorIndex;  // Returns or sets the color index that's applied to the foreground of the shading object.
@property MicrosoftWordMCoI foregroundPatternColorThemeIndex;  // Returns or sets the color that's applied to the foreground of the shading object. This color is applied to the dots and lines in the shading pattern.
@property MicrosoftWordE112 texture;  // Returns or sets the shading texture for the specified object.


@end

// Represents a contiguous area in a document. Each text range object is defined by a starting and ending character position.
@interface MicrosoftWordTextRange : MicrosoftWordBaseObject

- (SBElementArray *) characters;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) tables;
- (SBElementArray *) footnotes;
- (SBElementArray *) endnotes;
- (SBElementArray *) WordComments;
- (SBElementArray *) cells;
- (SBElementArray *) sections;
- (SBElementArray *) paragraphs;
- (SBElementArray *) fields;
- (SBElementArray *) formFields;
- (SBElementArray *) frames;
- (SBElementArray *) bookmarks;
- (SBElementArray *) revisions;
- (SBElementArray *) hyperlinkObjects;
- (SBElementArray *) subdocuments;
- (SBElementArray *) columns;
- (SBElementArray *) rows;
- (SBElementArray *) shapes;
- (SBElementArray *) readabilityStatistics;
- (SBElementArray *) grammaticalErrors;
- (SBElementArray *) spellingErrors;
- (SBElementArray *) inlineShapes;
- (SBElementArray *) mathObjects;
- (SBElementArray *) coauthLocks;
- (SBElementArray *) coauthUpdates;
- (SBElementArray *) conflicts;

@property BOOL bold;  // Returns or sets if the text associated with the text range is formatted as bold.
@property (readonly) NSInteger bookmarkId;  // Returns the number of the bookmark that encloses the beginning of the text range. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property MicrosoftWordE125 character_case;  // Synonym for case
@property MicrosoftWordE125 case1;  // Returns or sets a constant that represents the case of the text in the text range.
@property (copy, readonly) MicrosoftWordColumnOptions *columnOptions;
@property BOOL combineCharacters;
@property (copy) NSString *content;  // Returns or sets the text in the text range.
@property BOOL disableCharacterSpaceGrid;  // Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding text range object.
@property MicrosoftWordE114 emphasisMark;  // Returns or sets the emphasis mark for a character or designated character string.
@property (readonly) NSInteger endOfContent;  // Returns the ending character position of the text range.
@property (copy, readonly) MicrosoftWordEndnoteOptions *endnoteOptions;  // Returns the endnote options object for this text range.
@property (copy, readonly) MicrosoftWordFieldOptions *fieldOptions;
@property (copy, readonly) MicrosoftWordFind *findObject;  // Returns the find object associated with this text range.
@property double fitTextWidth;  // Returns or sets the width in which Microsoft Word fits the text in the text range.
@property (copy, readonly) MicrosoftWordFont *fontObject;  // Returns the font object associated with this text range.
@property (copy, readonly) MicrosoftWordFootnoteOptions *footnoteOptions;  // Returns the footnote options object for this text range.
@property (copy) MicrosoftWordTextRange *formattedText;  // Returns or sets a text range object that includes the formatted text in the text range.
@property BOOL grammarChecked;  // True if a grammar check has been run on the text range. False if some of the text range hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false.
@property MicrosoftWordE110 highlightColorIndex;  // Returns or sets the highlight color for the text range.
@property MicrosoftWordE309 horizontalInVertical;
@property (readonly) BOOL isEndOfRowMark;  // Returns true if the text range is collapsed and is located at the end-of-row mark in a table.
@property BOOL italic;  // Returns or sets if the text associated with the text range is formatted as italic.
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the text range object
@property MicrosoftWordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (copy, readonly) MicrosoftWordListFormat *listFormat;  // Returns the list format object associated with this text range.
@property (copy, readonly) MicrosoftWordTextRange *nextStoryRange;  // Returns a text range object that refers to the next story
@property BOOL noProofing;  // Returns or sets if the spelling and grammar checker should ignore documents based on this text range.
@property MicrosoftWordE270 orientation;  // Returns or sets the orientation of text in a text range when the text direction feature is enabled.
@property (copy) MicrosoftWordPageSetup *pageSetup;  // Returns or sets the page setup object associated with this text range.
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this text range.
@property (readonly) NSInteger previousBookmarkId;  // Returns the number of the last bookmark that starts before or at the same place as the text range, It returns zero if there's no corresponding bookmark.
@property (copy, readonly) MicrosoftWordRangeEndnoteOptions *rangeEndnoteOptions;
@property (copy, readonly) MicrosoftWordRangeFootnoteOptions *rangeFootnoteOptions;
@property (copy, readonly) MicrosoftWordRowOptions *rowOptions;
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with this text range.
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are included in the elements of this text range.
@property BOOL spellingChecked;  // True if a spelling check has been run on the text range. False if some of the text range hasn't been checked for spelling. To see if the document contains spelling errors, get the count of spelling errors for the text range.
@property (readonly) NSInteger startOfContent;  // Returns the starting character position of the text range.
@property (readonly) NSInteger storyLength;  // Returns the number of characters in the story that contains the text range.
@property (readonly) MicrosoftWordE160 storyType;  // Returns the story type for the text range.
@property MicrosoftWordE184 style;  // Returns or sets the Word style for this range.
@property BOOL subdocumentsExpanded;  // Returns or sets if the subdocuments in the specified text range are expanded.
@property MicrosoftWordE182 supplementalLanguageID;  // Returns or sets the language for the text range object
@property (copy) MicrosoftWordTextRetrievalMode *textRetrievalMode;  // Returns or sets the text retrieval object that controls how text is retrieved from this text range.
@property MicrosoftWordE308 twoLinesInOne;  // Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.  Read/write
@property MicrosoftWordE113 underline;  // Returns or sets the type of underline applied to the text range.

- (void) autoFormatTextRange;  // Automatically formats a text range.
- (double) calculateRange;  // Calculates a mathematical expression within a text range.
- (MicrosoftWordTextRange *) changeEndOfRangeBy:(MicrosoftWordE129)by extendType:(MicrosoftWordE249)extendType;  // Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) changeStartOfRangeBy:(MicrosoftWordE129)by extendType:(MicrosoftWordE249)extendType;  // Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.
- (BOOL) checkGrammar;  // Begins a grammar check for the text range.  Returns true if the text range had no errors
- (BOOL) checkSpellingCustomDictionary:(MicrosoftWordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(MicrosoftWordDictionary *)mainDictionary customDictionary2:(MicrosoftWordDictionary *)customDictionary2 customDictionary3:(MicrosoftWordDictionary *)customDictionary3 customDictionary4:(MicrosoftWordDictionary *)customDictionary4 customDictionary5:(MicrosoftWordDictionary *)customDictionary5 customDictionary6:(MicrosoftWordDictionary *)customDictionary6 customDictionary7:(MicrosoftWordDictionary *)customDictionary7 customDictionary8:(MicrosoftWordDictionary *)customDictionary8 customDictionary9:(MicrosoftWordDictionary *)customDictionary9 customDictionary10:(MicrosoftWordDictionary *)customDictionary10;  // Begins a spelling check for the text range.  Returns true if the text range had no errors
- (void) checkSynonyms;  // Displays the thesaurus dialog box, which lists alternative word choices, or synonyms, for the text in the text range.
- (MicrosoftWordTextRange *) collapseRangeDirection:(MicrosoftWordE132)direction;  // Collapses this text range to the starting or ending position and returns a new text range object. After a text range is collapsed, the starting and ending points are equal.
- (NSInteger) computeTextRangeStatisticsStatistic:(MicrosoftWordE155)statistic;  // Returns a statistic based on the contents of the specified text range.
- (MicrosoftWordTable *) convertToTableSeparator:(MicrosoftWordE177)separator numberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns initialColumnWidth:(NSInteger)initialColumnWidth tableFormat:(MicrosoftWordE180)tableFormat applyBorders:(BOOL)applyBorders applyShading:(BOOL)applyShading applyFont:(BOOL)applyFont applyColor:(BOOL)applyColor applyHeadingRows:(BOOL)applyHeadingRows applyLastRow:(BOOL)applyLastRow applyFirstColumn:(BOOL)applyFirstColumn applyLastColumn:(BOOL)applyLastColumn autoFit:(BOOL)autoFit autoFitBehavior:(MicrosoftWordE288)autoFitBehavior defaultTableBehavior:(MicrosoftWordE287)defaultTableBehavior;  // Converts text within a text range to a table.
- (void) copyAsPicture;// NS_RETURNS_NOT_RETAINED;  // Copies the content of the text range as a picture.
- (void) copyObject;// NS_RETURNS_NOT_RETAINED;  // Copies the content of the text range to the clipboard.
- (void) cutObject;  // Removes the content of the text range from the document and places it on the clipboard.
- (MicrosoftWordTextRange *) expandBy:(MicrosoftWordE129)by;  // Expands the specified range. This method returns the new range object or missing value if an error occurred.
- (NSString *) getRangeInformationInformationType:(MicrosoftWordE266)informationType;  // Returns requested information about the text range. 
- (MicrosoftWordTextRange *) goToNextWhat:(MicrosoftWordE130)what;  // Returns a text range object that refers to the start position of the next item or location specified by the what argument.
- (MicrosoftWordTextRange *) goToPreviousWhat:(MicrosoftWordE130)what;  // Returns a text range object that refers to the start position of the previous item or location specified by the what argument.
- (BOOL) inRangeTextRange:(MicrosoftWordTextRange *)textRange;  // Returns true if the text range to which the method is applied is contained in the range specified by the text range argument.
- (BOOL) inStoryTextRange:(MicrosoftWordTextRange *)textRange;  // Returns true if the text range to which the method is applied is in the same story as the range specified by the text range argument.
- (BOOL) isEquivalentTextRange:(MicrosoftWordTextRange *)textRange;  // True if the selection or range to which this method is applied is equal to the range specified by the text range argument. This method compares the starting and ending character positions, as well as the story type.
- (void) modifyEnclosureEnclosureStyle:(MicrosoftWordE277)enclosureStyle enclosureType:(MicrosoftWordE276)enclosureType enclosedText:(NSString *)enclosedText;  // Adds, modifies, or removes an enclosure around the specified character or characters.
- (MicrosoftWordTextRange *) moveEndOfRangeBy:(MicrosoftWordE129)by count:(NSInteger)count;  // Moves the ending character position of the range.  This method returns the new text range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeBy:(MicrosoftWordE129)by count:(NSInteger)count;  // Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeEndUntilCharacters:(NSString *)characters count:(MicrosoftWordE294)count;  // Moves the end position of the specified range until any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeEndWhileCharacters:(NSString *)characters count:(MicrosoftWordE294)count;  // Moves the ending character position of a the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeStartUntilCharacters:(NSString *)characters count:(MicrosoftWordE294)count;  // Moves the start position of the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeStartWhileCharacters:(NSString *)characters count:(MicrosoftWordE294)count;  // Moves the start position of the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeUntilCharacters:(NSString *)characters count:(MicrosoftWordE294)count;  // Moves the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveRangeWhileCharacters:(NSString *)characters count:(MicrosoftWordE294)count;  // Moves the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) moveStartOfRangeBy:(MicrosoftWordE129)by count:(NSInteger)count;  // Moves the starting character position of the range. This method returns the new text range object or missing value if an error occurred.
- (MicrosoftWordTextRange *) navigateTo:(MicrosoftWordE130)to position:(MicrosoftWordE131)position count:(NSInteger)count name:(NSString *)name;  // Returns a text range object that represents the start position of the specified item, such as a page, bookmark, or field.
- (MicrosoftWordTextRange *) nextRangeBy:(MicrosoftWordE129)by count:(NSInteger)count;  // Returns a text range object relative to the specified text range.
- (MicrosoftWordTextRange *) nextSubdocument;  // Returns a new text range object to the next subdocument. If there isn't another subdocument, an error occurs.
- (void) pasteAndFormatType:(MicrosoftWordE293)type;  // Pastes the selected table cells and formats them as specified.
- (void) pasteAppendTable;  // Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.
- (void) pasteAsNestedTable;  // Pastes a cell or group of cells as a nested table into the text range.
- (void) pasteExcelTableLinkedToExcel:(BOOL)linkedToExcel wordFormatting:(BOOL)wordFormatting RTF:(BOOL)RTF;  // Pastes and formats a Microsoft Excel table.
- (void) pasteObject;  // Inserts the contents of the clipboard at the specified text range.
- (void) pasteSpecialLink:(BOOL)link placement:(MicrosoftWordE243)placement displayAsIcon:(BOOL)displayAsIcon dataType:(MicrosoftWordE251)dataType iconLabel:(NSString *)iconLabel;  // Inserts the contents of the clipboard. Unlike with the paste method, with paste special you can control the format of the pasted information and optionally establish a link to the source file - for example, a Microsoft Excel worksheet.
- (void) phoneticGuideText:(NSString *)text alignmentType:(MicrosoftWordE310)alignmentType raise:(NSInteger)raise fontSize:(NSInteger)fontSize fontName:(NSString *)fontName;
- (MicrosoftWordTextRange *) previousRangeBy:(MicrosoftWordE129)by count:(NSInteger)count;  // Returns a text range object relative to the specified text range.
- (MicrosoftWordTextRange *) previousSubdocument;  // Returns a new text range object to the previous subdocument. If there isn't another subdocument, an error occurs.
- (void) relocateDirection:(MicrosoftWordE224)direction;  // In outline view, moves the paragraphs within the text range after the next visible paragraph or before the previous visible paragraph. Body text moves with a heading only if the body text is collapsed in outline view or if it's part of the text range.
- (MicrosoftWordTextRange *) setRangeStart:(NSInteger)start end:(NSInteger)end;  // Returns a text range object by using the specified starting and ending character positions.
- (void) sortExcludeHeader:(BOOL)excludeHeader fieldNumber:(NSInteger)fieldNumber sortFieldType:(MicrosoftWordE178)sortFieldType sortOrder:(MicrosoftWordE179)sortOrder fieldNumberTwo:(NSInteger)fieldNumberTwo sortFieldTypeTwo:(MicrosoftWordE178)sortFieldTypeTwo sortOrderTwo:(MicrosoftWordE179)sortOrderTwo fieldNumberThree:(NSInteger)fieldNumberThree sortFieldTypeThree:(MicrosoftWordE178)sortFieldTypeThree sortOrderThree:(MicrosoftWordE179)sortOrderThree sortColumn:(BOOL)sortColumn separator:(MicrosoftWordE176)separator caseSensitive:(BOOL)caseSensitive languageID:(MicrosoftWordE182)languageID;  // Sorts the paragraphs in the specified text range.
- (void) sortAscending;  // Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) sortDescending;  // Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (NSDictionary *) textRangeSpellingSuggestionsCustomDictionary:(MicrosoftWordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(MicrosoftWordDictionary *)mainDictionary suggestionMode:(MicrosoftWordE256)suggestionMode customDictionary2:(MicrosoftWordDictionary *)customDictionary2 customDictionary3:(MicrosoftWordDictionary *)customDictionary3 customDictionary4:(MicrosoftWordDictionary *)customDictionary4 customDictionary5:(MicrosoftWordDictionary *)customDictionary5 customDictionary6:(MicrosoftWordDictionary *)customDictionary6 customDictionary7:(MicrosoftWordDictionary *)customDictionary7 customDictionary8:(MicrosoftWordDictionary *)customDictionary8 customDictionary9:(MicrosoftWordDictionary *)customDictionary9 customDictionary10:(MicrosoftWordDictionary *)customDictionary10;  // Returns a record that contains the spelling error type and the list of words suggested as replacements for the first word in the specified range

@end

@interface MicrosoftWordCharacter : MicrosoftWordTextRange


@end

@interface MicrosoftWordGrammaticalError : MicrosoftWordTextRange


@end

@interface MicrosoftWordSentence : MicrosoftWordTextRange


@end

@interface MicrosoftWordSpellingError : MicrosoftWordTextRange


@end

@interface MicrosoftWordStoryRange : MicrosoftWordTextRange


@end

@interface MicrosoftWordWord : MicrosoftWordTextRange


@end



/*
 * Proofing Suite
 */

// Represents a single autocorrect entry.
@interface MicrosoftWordAutocorrectEntry : MicrosoftWordBaseObject

@property (copy) NSString *autocorrectValue;  // Returns or sets the value of the auto correct entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of the auto correct entry.
@property (readonly) BOOL richText;  // Returns if formatting is stored with the autocorrect entry replacement text.

- (void) applyCorrectionToRange:(MicrosoftWordTextRange *)toRange;  // Replaces a range with the value of the specified autocorrect entry.

@end

// Represents the autocorrect functionality in Word.
@interface MicrosoftWordAutocorrect : MicrosoftWordBaseObject

- (SBElementArray *) autocorrectEntries;
- (SBElementArray *) firstLetterExceptions;
- (SBElementArray *) twoInitialCapsExceptions;
- (SBElementArray *) otherCorrectionsExceptions;

@property BOOL correctDays;  // Returns or sets if Word automatically capitalizes the first letter of days of the week.
@property BOOL correctInitialCaps;  // Returns or sets if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, WOrd is corrected to Word.
@property BOOL correctSentenceCaps;  // Returns or sets if Word automatically capitalizes the first letter in each sentence.
@property BOOL correctTableCaps;  // Returns or sets if Word automatically capitalizes the first letter in each table cell.
@property BOOL firstLetterAutoAdd;  // Returns or sets if Word automatically adds abbreviations to the list of autocorrect first letter exceptions.
@property BOOL otherCorrectionsAutoAdd;  // Returns or sets if Microsoft Word automatically adds words to the list of other autocorrect exceptions. Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct.
@property BOOL replaceText;  // Returns or sets if Microsoft Word automatically replaces specified text with entries from the autocorrect list.
@property BOOL replaceTextFromSpellingChecker;  // Returns or sets if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types.
@property BOOL showAutocorrectSmartButton;  // Returns or sets if Word shows the AutoCorrect smart button which allows you to review the correction when an automatic correction occurs.
@property BOOL turnOnAutocorrect;  // Returns or sets if Word automatically corrects spelling and formatting as you type.
@property BOOL twoInitialCapsAutoAdd;  // Returns or sets if Microsoft Word automatically adds words to the list of autocorrect initial caps exceptions.


@end

// Represents a dictionary.
@interface MicrosoftWordDictionary : MicrosoftWordBaseObject

@property (readonly) MicrosoftWordE255 dictionaryType;  // Returns the dictionary type.
@property MicrosoftWordE182 languageID;  // Returns or sets the language for the dictionary object
@property BOOL languageSpecific;  // Returns or sets if the custom dictionary is to be used only with text formatted for a specific language.
@property (copy, readonly) NSString *name;  // Returns or sets the name of this dictionary object.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified dictionary.
@property (readonly) BOOL readOnly;  // Returns true if the specified dictionary cannot be changed.


@end

// Represents an abbreviation excluded from automatic correction.
@interface MicrosoftWordFirstLetterException : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of this first letter exception object.


@end

// Represents a language used for proofing or formatting in Microsoft Word.
@interface MicrosoftWordLanguage : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordDictionary *activeGrammarDictionary;  // Returns a dictionary object that represents the active grammar dictionary for the specified language
@property (copy, readonly) MicrosoftWordDictionary *activeHyphenationDictionary;  // Returns a dictionary object that represents the active hyphenation dictionary for the specified language.
@property (copy, readonly) MicrosoftWordDictionary *activeSpellingDictionary;  // Returns a dictionary object that represents the active spelling dictionary for the specified language.
@property (copy, readonly) MicrosoftWordDictionary *activeThesaurusDictionary;  // Returns a dictionary object that represents the active thesaurus dictionary for the specified language.
@property (copy) NSString *defaultWritingStyle;  // Returns or sets the default writing style used by the grammar checker for the specified language. The name of the writing style is the localized name for the specified language.
@property (readonly) MicrosoftWordE182 languageID;  // Returns an enumeration that identifies the specified language.
@property (copy, readonly) NSString *name;  // Returns the name of the language
@property (copy, readonly) NSString *nameLocal;  // Returns the name of a proofing tool language in the language of the user.
@property MicrosoftWordE255 spellingDictionaryType;  // Returns or sets the proofing tool type
@property (copy, readonly) NSArray *writingStyleList;  // Returns a list of strings that contains the names of all writing styles available for the specified language.


@end

// Represents a single auto correct exception.
@interface MicrosoftWordOtherCorrectionsException : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of this other corrections exception object.


@end

// Represents one of the readability statistics for a document or range.
@interface MicrosoftWordReadabilityStatistic : MicrosoftWordBaseObject

@property (copy, readonly) NSString *name;  // The name of this readability statistic object.
@property (readonly) double readabilityValue;  // The value of this readability statistic object.


@end

// Represents the information about synonyms, antonyms, related words, or related expressions for the specified range or a given string.
@interface MicrosoftWordSynonymInfo : MicrosoftWordBaseObject

@property (copy, readonly) NSArray *antonyms;  // Returns a list of antonyms for the word or phrase.
@property (copy, readonly) NSString *context;  // Returns the word or phrase that was looked up by the thesaurus.
@property (readonly) BOOL found;  // Returns true if the thesaurus finds synonyms, antonyms, related words, or related expressions for the word or phrase.
@property (readonly) NSInteger meaningCount;  // Returns the number of entries in the list of meanings found in the thesaurus for the word or phrase. Returns zero if no meanings were found.
@property (copy, readonly) NSArray *meanings;  // Returns the list of meanings for the word or phrase.
@property (copy, readonly) NSArray *partOfSpeech;  // Returns a list of the parts of speech corresponding to the meanings found for the word or phrase looked up in the thesaurus.
@property (copy, readonly) NSArray *relatedExpressions;  // Returns a list of expressions related to the specified word or phrase. 
@property (copy, readonly) NSArray *relatedWords;  // Returns a list of words related to the specified word or phrase.

- (NSArray *) getSynonymListForItemToCheck:(NSString *)itemToCheck;  // Get the list of synonyms for a particular word
- (NSArray *) getSynonymListFromMeaningIndex:(NSInteger)meaningIndex;  // Get the list of synonyms using an index into the list of meanings

@end

// Represents a single initial-capital autocorrect exception.
@interface MicrosoftWordTwoInitialCapsException : MicrosoftWordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of this two initial caps exception object.


@end



/*
 * Table Suite
 */

// Represents a single table cell.
@interface MicrosoftWordCell : MicrosoftWordBaseObject

- (SBElementArray *) tables;

@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordColumn *column;  // Returns the column object that contains this cell object.
@property (readonly) NSInteger columnIndex;  // Returns the number of the column that contains the specified cell.
@property BOOL fitText;  // Returns or sets if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width.
@property double height;  // Returns or sets the height of the object.
@property MicrosoftWordE133 heightRule;  // Returns or sets the rule for determining the height of the specified cells.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified cell.
@property (copy, readonly) MicrosoftWordCell *nextCell;  // Returns the next cell object
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified cell.
@property MicrosoftWordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified column.
@property (copy, readonly) MicrosoftWordCell *previousCell;  // Returns the previous cell object
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordRow *row;  // Returns the row object that contains this cell object.
@property (readonly) NSInteger rowIndex;  // Returns the number of the row that contains the specified cell.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the cell object.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the cell object.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property MicrosoftWordE147 verticalAlignment;  // Returns or sets the vertical alignment of text in one or more cells of a table.
@property double width;  // Returns or sets the width of the object.
@property BOOL wordWrap;  // Returns or set  if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same.

- (void) autoSum;  // Inserts an = Formula field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.
- (void) formulaFormulaString:(NSString *)formulaString numberFormatString:(NSString *)numberFormatString;  // Inserts an = Formula field that contains the specified formula into a table cell.
- (void) mergeCellWith:(MicrosoftWordCell *)with;  // Merges the specified table cell with another cell. The result is a single table cell.
- (void) splitCellNumberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns;  // Splits a single table cell into multiple cells.

@end

// Represents options that can be set for columns.
@interface MicrosoftWordColumnOptions : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double defaultWidth;  // Returns or sets the default width of columns.
@property (readonly) NSInteger nestingLevel;
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified columns. 
@property MicrosoftWordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified columns. 
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with columns.

- (void) distributeWidth;  // Adjusts the width of the specified columns or cells so that they're equal.

@end

// Represents a single table column.
@interface MicrosoftWordColumn : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this column object.
@property (readonly) NSInteger columnIndex;  // Returns the index for the position of the object in its container element list.
@property (readonly) BOOL isFirst;  // Returns if the specified column is the first one in the table.
@property (readonly) BOOL isLast;  // Returns if the specified column is the last one in the table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified column.
@property (copy, readonly) MicrosoftWordColumn *nextColumn;  // Returns the next column object
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified column.
@property MicrosoftWordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified column.
@property (copy, readonly) MicrosoftWordColumn *previousColumn;  // Returns the previous column object
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the column object.
@property double width;  // Returns or sets the width of the object.


@end

@interface MicrosoftWordCondition : MicrosoftWordBaseObject

@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns the border options object associated with this condition.
@property double bottomPadding;
@property (copy, readonly) MicrosoftWordFont *fontObject;
@property double leftPadding;
@property (copy) MicrosoftWordParagraphFormat *paragraphFormat;
@property double rightPadding;
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with this condition.
@property double topPadding;


@end

// Represents options that can be set for rows.
@interface MicrosoftWordRowOptions : MicrosoftWordBaseObject

@property MicrosoftWordE144 alignment;  // Returns or sets a constant that represents the alignment for rows.
@property BOOL allowBreakAcrossPages;  // Returns or sets if the text in a table row or rows are allowed to split across a page break.
@property BOOL allowOverlap;  // Returns or sets a value that specifies whether the specified rows can overlap other rows.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double distanceBottom;  // Returns or sets the distance in points between the document text and the bottom edge of the specified table. This property doesn't have any effect if wrap around text is false. 
@property double distanceLeft;  // Returns or sets the distance in points between the document text and the left edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property double distanceRight;  // Returns or sets the distance in points between the document text and the right edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property double distanceTop;  // Returns or sets the distance in points between the document text and the top edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property BOOL headingFormat;  // Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page.
@property double height;  // Returns or sets the height of the object.
@property MicrosoftWordE133 heightRule;  // Returns or sets the rule for determining the height of the specified rows.
@property double horizontalPosition;  // Returns or sets the horizontal distance between the edge of the rows.
@property (readonly) NSInteger nestingLevel;
@property MicrosoftWordE236 relativeHorizontalPosition;  // Specifies to what the horizontal position of a group of rows is relative.
@property MicrosoftWordE237 relativeVerticalPosition;  // Specifies to what the vertical position of a group of rows is relative.
@property double rowLeftIndent;  // Returns or sets the left indent in points for the specified rows.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with rows.
@property double spaceBetweenColumns;  // Returns or sets the distance in points between text in adjacent columns of the specified row or rows.
@property double verticalPosition;  // Returns or sets the vertical distance between the edge of the rows.
@property BOOL wrapAroundText;  // Returns or sets whether text should wrap around the specified rows. 

- (void) distributeRowHeight;  // Adjusts the height of the specified rows or cells so that they're equal.

@end

// Represents a row in a table.
@interface MicrosoftWordRow : MicrosoftWordBaseObject

- (SBElementArray *) cells;

@property MicrosoftWordE144 alignment;  // Returns or sets a constant that represents the alignment for the specified row.
@property BOOL allowBreakAcrossPages;  // Returns or sets if the text in a row or rows are allowed to split across a page break.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property BOOL headingFormat;  // Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page.
@property double height;  // Returns or sets the height of the object.
@property MicrosoftWordE133 heightRule;  // Returns or sets the rule for determining the height of the specified rows.
@property (readonly) BOOL isFirst;  // Returns if the specified row is the first one in the table.
@property (readonly) BOOL isLast;  // Returns if the specified row is the last one in the table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified row.
@property (copy, readonly) MicrosoftWordRow *nextRow;  // Returns the next row object
@property (copy, readonly) MicrosoftWordRow *previousRow;  // Returns the previous row object
@property (readonly) NSInteger rowIndex;  // Returns the index for the position of the object in its container element list.
@property double rowLeftIndent;  // Returns or sets the left indent in points for the specified rows.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the row object.
@property double spaceBetweenColumns;  // Returns or sets the distance in points between text in adjacent columns of the specified row or rows.  
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the row object.


@end

@interface MicrosoftWordTableStyle : MicrosoftWordBaseObject

@property MicrosoftWordE144 alignment;  // Returns or sets a constant that represents the alignment for the specified row.
@property BOOL allowBreakAccrossPage;
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property (copy, readonly) MicrosoftWordCondition *bottomLeftCellCondition;  // Returns the conditional style for the bottom left cell.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordCondition *bottomRightCellCondition;  // Returns the conditional style for the bottom right cell.
@property NSInteger columnStripe;
@property (copy, readonly) MicrosoftWordCondition *evenColumnCondition;  // Returns the conditional style for even column bands.
@property (copy, readonly) MicrosoftWordCondition *evenRowCondition;  // Returns the conditional style for even row bands.
@property (copy, readonly) MicrosoftWordCondition *firstColumnCondition;  // Returns the conditional style for the first column.
@property (copy, readonly) MicrosoftWordCondition *firstRowCondition;  // Returns the conditional style for the first row.
@property (copy, readonly) MicrosoftWordCondition *lastColumnCondition;  // Returns the conditional style for the last column.
@property (copy, readonly) MicrosoftWordCondition *lastRowCondition;  // Returns the conditional style for the last row.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordCondition *oddColumnCondition;  // Returns the conditional style for odd column bands.
@property (copy, readonly) MicrosoftWordCondition *oddRowCondition;  // Returns the conditional style for odd row bands.
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property double rowLeftIndent;  // Returns or sets the left indent in points for this table style.
@property NSInteger rowStripe;
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with this table style.
@property double spacing;  // Returns or sets the spacing between the cells in a table.
@property (copy, readonly) MicrosoftWordCondition *topLeftCellCondition;  // Returns the conditional style for the top left cell.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordCondition *topRightCellCondition;  // Returns the conditional style for the top right cell.


@end

// Represents a single table.
@interface MicrosoftWordTable : MicrosoftWordBaseObject

- (SBElementArray *) columns;
- (SBElementArray *) rows;
- (SBElementArray *) tables;

@property BOOL allowAutoFit;  // Returns or sets if Microsoft Word will automatically resize cells in a table to fit their contents.
@property BOOL allowPageBreaks;  // Returns or sets if Microsoft Word allows to break the specified table across pages.
@property BOOL applyStyleFirstColumn;
@property BOOL applyStyleHeadingRows;
@property BOOL applyStyleLastColumn;
@property BOOL applyStyleLastRow;
@property (readonly) MicrosoftWordE180 autoFormatType;  // Returns the type of automatic formatting that's been applied to the specified table.
@property (copy, readonly) MicrosoftWordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordColumnOptions *columnOptions;  // Returns the column options object associated with this table object.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified table.
@property (readonly) NSInteger numberOfColumns;  // Returns the number of columns in this table
@property (readonly) NSInteger numberOfRows;  // Returns the number of rows in this table
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified table.
@property MicrosoftWordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified table. 
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property (copy, readonly) MicrosoftWordRowOptions *rowOptions;  // Returns the row options object associated with this table object.
@property (copy, readonly) MicrosoftWordShading *shading;  // Returns the shading object associated with the table object.
@property double spacing;  // Returns or sets the spacing between the cells in a table.
@property MicrosoftWordE184 style;  // Returns or sets the Word style associated with the table object.
@property (copy, readonly) MicrosoftWordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the table object.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property (readonly) BOOL uniform;  // Returns if all the rows in a table have the same number of columns.

- (void) autoFitBehaviorBehavior:(MicrosoftWordE288)behavior;  // Determines how Microsoft Word resizes a table when the autofit feature is used. Word can resize the table based on the content of the table cells or the width of the document window.
- (void) autoFormatTableTableFormat:(MicrosoftWordE180)tableFormat applyBorders:(BOOL)applyBorders applyShading:(BOOL)applyShading applyFont:(BOOL)applyFont applyColor:(BOOL)applyColor applyHeadingRows:(BOOL)applyHeadingRows applyLastRow:(BOOL)applyLastRow applyFirstColumn:(BOOL)applyFirstColumn applyLastColumn:(BOOL)applyLastColumn autoFit:(BOOL)autoFit;  // Applies a predefined look to a table.
- (MicrosoftWordTextRange *) convertToTextSeparator:(MicrosoftWordE177)separator nestedTables:(BOOL)nestedTables;  // Converts a table to text and returns a text range object that represents the delimited text.
- (MicrosoftWordCell *) getCellFromTableRow:(NSInteger)row column:(NSInteger)column;  // Returns a cell object that represents a cell in a table.
- (MicrosoftWordTable *) splitTableRow:(NSInteger)row;  // Inserts an empty paragraph immediately above the specified row in the table, and returns a table object that contains both the specified row and the rows that follow it.
- (void) updateAutoFormat;  // Updates the table with the characteristics of a predefined table format. For example, if you apply a table format with auto format and then insert rows and columns, the table may no longer match the predefined look.

@end

