# Excel-Spreadsheet-Utils
Utility class for creating spreadsheets from data stored in xml files under the res/xml Android project resources folder. The spreadsheets are stored in external storage under the folder ResourceSpreadsheets.

Primary use is for creating a spreadsheet with to consolidate similarly named fields listed in different Android resource files into one Excel spreadsheet. The source files should be placed in the project res/xml folder. Corresponding data from each file will be packaged into a spreadsheet under column names specified by the user.

As an example, an Android application can have strings.xml files for different locales(res/values/strings/strings.xml, res/values-es/strings/strings.xml, res/values-fr/strings.xml). If we would like a spreadsheet listing all of the translations side by side, under columns named English, Spanish and French, the xml files will copied, renamed and placed into the res/xml/ folder:

res/values/strings/strings.xml -> res/xml/strings-default.xml
res/values-es/strings/strings.xml -> res/xml/strings-spanish.xml
res/values-fr/strings/strings.xml -> res/xml/strings-french.xml

In our source code, the spreadsheet creation will then proceed as follows.

        List<SpreadsheetUtils.TranslationSet> translationSetList = new ArrayList<>();
        translationSetList.add(new SpreadsheetUtils.TranslationSet(R.xml.strings-default, "English"));
        translationSetList.add(new SpreadsheetUtils.TranslationSet(R.xml.strings-spanish, "Spanish"));
        translationSetList.add(new SpreadsheetUtils.TranslationSet(R.xml.strings-french, "French"));
        final String outputFileName = "translations.xls";
        new SpreadsheetUtils.StringResourcesSpreadsheetTask(InstrumentationRegistry.getTargetContext(),
                outputFileName,
                translationSetList)
                .execute();


    
