
package com.helper.spreadsheet;

import android.content.Context;
import android.content.res.XmlResourceParser;
import android.os.AsyncTask;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.annotation.Nullable;
import android.support.annotation.XmlRes;
import android.util.Log;

import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public final class SpreadsheetUtils {
    private static final String TAG = SpreadsheetUtils.class.getName();
    private static final int INVALID = -1;
    private static final String RESOURCE_SPREADSHEET_FOLDER = "/ResourceSpreadsheets";
    private static final String DATA_SHEET_NAME = "Data";
    private static final String ATTRIBUTE_NAME = "name";

    // region Resource Value Extraction
    void buildStringResourcesSpreadsheet(@NonNull Context context,@NonNull String filename, @NonNull List<TranslationSet> translationSets) {
        if (translationSets != null && translationSets.size() > 0) {
            List<Map<String, String>> maps = new ArrayList<>(translationSets.size());
            for (TranslationSet set : translationSets) {
                maps.add(createResourceDataMap(context, set.resId, maps.size() > 0 ?
                        maps.get(0).keySet() : null));
            }
            createSpreadsheet(filename,translationSets, maps);
        }
    }

    private Map<String, String> createResourceDataMap(@NonNull Context context, @XmlRes int xmlResId, @Nullable
            Set<String> tagSet) {
        boolean skip = false;
        String currentTag = null;

        Map<String, String> stringValuesMap = new HashMap<>();
        try {
            XmlResourceParser parser = context.getResources().getXml(xmlResId);
            int eventType = parser.getEventType();
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    int nameAttributeIndex = getNameAttributeIndex(parser);
                    if (nameAttributeIndex == INVALID) {
                        skip = true;
                    } else if (tagSet != null && !tagSet.contains(parser.getAttributeValue(nameAttributeIndex))) {
                        skip = true;
                    } else {
                        currentTag = parser.getAttributeValue(nameAttributeIndex);
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    currentTag = null;
                    skip = false;
                } else if (eventType == XmlPullParser.TEXT) {
                    if (!skip && currentTag != null) {
                        stringValuesMap.put(currentTag, parser.getText());
                    }
                }
                eventType = parser.next();
            }
            parser.close();
        } catch (IOException e) {
            Log.d(TAG, e.getMessage());
        } catch (XmlPullParserException e) {
            Log.d(TAG, e.getMessage());
        }

        return stringValuesMap;
    }

    private int getNameAttributeIndex(@NonNull XmlResourceParser parser) {
        if (parser.getAttributeCount() > 0) {
            for (int i = 0; i < parser.getAttributeCount(); i++) {
                if (parser.getAttributeName(i).equals(ATTRIBUTE_NAME)) {
                    return i;
                }
            }
        }

        return INVALID;
    }

    // endregion

    // region Spreadsheet Creation
    private void createSpreadsheet(@NonNull String filename,@NonNull List<TranslationSet> translationSets,
            List<Map<String, String>> maps) {
        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setUseTemporaryFileDuringWrite(true);

        File sdCard = Environment.getExternalStorageDirectory();
        File dir = new File(sdCard.getAbsolutePath() + RESOURCE_SPREADSHEET_FOLDER);
        dir.mkdirs();
        File wbfile = new File(dir, filename);

        WritableWorkbook wb = null;
        try {
            wbfile.createNewFile();
            wb = Workbook.createWorkbook(wbfile, wbSettings);

            if (maps.size() > 0) {
                wb.createSheet(DATA_SHEET_NAME, 0);
                WritableSheet sheet = wb.getSheet(0);

                int column = 0;
                for (int i = 0; i < translationSets.size(); i++) {
                    writeCell(column, 0, translationSets.get(column).columnName, true, sheet);
                }
                column = 0;
                int row = 1;
                Map<String, String> firstMap = maps.get(0);
                for (Map.Entry<String, String> entry : firstMap.entrySet()) {
                    writeCell(column, row, entry.getValue(), false, sheet);
                    for (int i = 1; i < maps.size(); i++) {
                        column++;
                        Map<String, String> nextMap = maps.get(i);
                        if (nextMap.containsKey(entry.getKey())) {
                            writeCell(column, row, nextMap.get(entry.getKey()), false, sheet);
                        }
                    }
                    column = 0;
                    row++;
                }
            }

            wb.write();
        } catch (IOException ex) {
            Log.d(TAG, ex.getMessage());
        } catch (RowsExceededException e) {
            Log.d(TAG, e.getMessage());
        } catch (WriteException e) {
            e.printStackTrace();
        } finally {
            try {
                if (wb != null) {
                    wb.close();
                }
            } catch (IOException e) {
                Log.d(TAG, e.getMessage());
            } catch (WriteException e) {
                Log.d(TAG, e.getMessage());
            }
        }
    }

    /**
     * @param columnPosition - column to place new cell in
     * @param rowPosition - row to place new cell in
     * @param contents - string value to place in cell
     * @param headerCell - whether to give this cell special formatting
     * @param sheet - WritableSheet to place cell in
     */
    private void writeCell(int columnPosition, int rowPosition, String contents, boolean headerCell,
            WritableSheet sheet) throws RowsExceededException, WriteException {
        //create a new cell with contents at position
        Label newCell = new Label(columnPosition, rowPosition, contents);
        if (headerCell) {
            WritableFont headerFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat headerFormat = new WritableCellFormat(headerFont);
            headerFormat.setAlignment(Alignment.CENTRE);
            newCell.setCellFormat(headerFormat);
        }

        sheet.addCell(newCell);
    }
    // endregion

    /**
     * Task for creating an Excel spreadsheet to consolidate all of the string translations.
     * <p>
     * The string resource files for each locale should be placed inside the res/xml folder and renamed according
     * to their locale since they all share the same name (strings.xml) in thier respect res/values-XX/strings folder.
     */
    public static class StringResourcesSpreadsheetTask extends AsyncTask<Void, Void, Void> {
        private final Context mContext;
        private final List<TranslationSet> mTranslationSets;
        private final String mFilename;

        /**
         * @param filename File name for the output spreadsheet. Should have .xls extension.
         * @param translationSets List of {@link TranslationSet}s identifying the string resource file and the column name
         * under which the values should be placed. The {@link TranslationSet} corresponding to the primary or default
         * string file should be supplied first in the list. The first entry is used to determine what set of attributes
         */
        public StringResourcesSpreadsheetTask(@NonNull Context context, @NonNull String filename,
                List<TranslationSet> translationSets) {
            mContext = context;
            mFilename = filename;
            mTranslationSets = translationSets;
        }

        @Override
        protected Void doInBackground(Void... voids) {
            new SpreadsheetUtils().buildStringResourcesSpreadsheet(mContext,mFilename, mTranslationSets);
            return null;
        }
    }

    /**
     * Points to an identifier reference a strings resource file and the column name that the values taken from that
     * file should be placed under in the spreadsheet.
     */
    public static class TranslationSet {
        @XmlRes
        private final int resId;
        private final String columnName;

        /**
         * @param stringsResId Resource id of the xml file containing the string resources. This file should be placed in the res/xml
         * folder.
         * @param columnName Name of the column in the spreadsheet to place the string resource values under.
         */
        public TranslationSet(@XmlRes int stringsResId, @NonNull String columnName) {
            resId = stringsResId;
            this.columnName = columnName;
        }
    }
}
