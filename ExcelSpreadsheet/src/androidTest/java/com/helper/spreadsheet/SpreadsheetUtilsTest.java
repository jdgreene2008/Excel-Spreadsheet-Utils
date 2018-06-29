
package com.helper.spreadsheet;

import android.support.test.InstrumentationRegistry;

import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class SpreadsheetUtilsTest {

    @Test
    public void buildStringResourcesSpreadsheet() {
        // Sample test using

        InstrumentationRegistry.getTargetContext();
        List<SpreadsheetUtils.TranslationSet> translationSetList = new ArrayList<>();
        translationSetList.add(new SpreadsheetUtils.TranslationSet(R.xml.strings, "English"));
        translationSetList.add(new SpreadsheetUtils.TranslationSet(R.xml.strings_es, "Spanish"));
        translationSetList.add(new SpreadsheetUtils.TranslationSet(R.xml.strings_fr, "French"));
        final String outputFileName = "translations.xls";
        new SpreadsheetUtils.StringResourcesSpreadsheetTask(InstrumentationRegistry.getTargetContext(),
                outputFileName,
                translationSetList)
                .execute();
    }
}
