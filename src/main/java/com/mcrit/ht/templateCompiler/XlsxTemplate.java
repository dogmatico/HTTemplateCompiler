/*
 * COPYRIGHT (c) 2016 MCRIT - Cristian Lorenzo Martinez <cristian.lorenzo.martinez@gmail.com>
 * MIT License
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:

 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.

 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

package com.mcrit.ht.templateCompiler;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;
import java.util.Arrays;
import java.util.stream.IntStream;
import java.util.Base64;
import javax.json.Json;
import javax.json.JsonArray;
import javax.json.JsonObject;
import javax.json.JsonReader;
import javax.json.JsonValue;
import org.apache.poi.ss.formula.IStabilityClassifier;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 *
 * @author Cristian Lorenzo <cristian.lorenzo.martinez@gmail.com>
 */
public class XlsxTemplate {
    private static XSSFWorkbook workbook;
    private static final boolean debug = false;

    public XlsxTemplate (String fileName) throws FileNotFoundException, IOException {
        FileInputStream fis = new FileInputStream(fileName);
        try {
	    workbook = new XSSFWorkbook(fis);
	} finally {
	    fis.close();
	}
    }
    
    private void populateTextSheet(JsonArray target, JsonArray data) {
        XSSFSheet sheet = workbook.getSheetAt(target.getInt(0));

        for (int j = 0, ln2 = data.size(); j < ln2; j++) {
            XSSFRow row = sheet.getRow(j + target.getJsonArray(1).getInt(1));
            if (row == null) {
                sheet.createRow(j + target.getJsonArray(1).getInt(1));
                row = sheet.getRow(j + target.getJsonArray(1).getInt(1));
            }
            for (int k = 0, ln3 = data.getJsonArray(j).size(); k < ln3; k++) {
                XSSFCell cell = row.getCell(k + target.getJsonArray(1).getInt(0), XSSFRow.CREATE_NULL_AS_BLANK);

                switch (data.getJsonArray(j).get(k).getValueType()) {
                    case STRING:
                        cell.setCellValue(data.getJsonArray(j).getString(k));
                        break;
                    case NUMBER:
                        cell.setCellValue(data.getJsonArray(j).getJsonNumber(k).doubleValue());
                        break;
                }
            }
        } 
    }
    
    /**
     * Adds and resize a image encoded in base64.
     * 
     * @param target JSON Array with the structure [
     *      (int) Target sheet number,
     *      [(int) Col1, (int) Row 1],
     *      [(int) Col2, (int) Row 2]
     *  ]
     * @param data JSON Array with the structure [
     *      (String) Image type,
     *      (String) Base64 encoded image
     *  ]
     */
    private void addImageBase64ToSheet(JsonArray target, JsonArray data) {
       int imageType;
       switch (data.getString(0).toLowerCase()) {
           case "dib":
               imageType = Workbook.PICTURE_TYPE_DIB;
               break;
            case "emf":
               imageType = Workbook.PICTURE_TYPE_EMF;
               break;
            case "jpeg":
            case "jpg":
               imageType = Workbook.PICTURE_TYPE_JPEG;
               break;
            case "pict":
               imageType = Workbook.PICTURE_TYPE_PICT;
               break;
            case "png":
               imageType = Workbook.PICTURE_TYPE_PNG;
               break;
            case "wmf":
               imageType = Workbook.PICTURE_TYPE_WMF;
               break;
           default:
               throw new RuntimeException("The provided format type is not supported.");
        }

        final byte[] decodedImg = Base64.getDecoder().decode(data.getString(1));
        final int pictureIndex = workbook.addPicture(decodedImg, imageType);

        final CreationHelper helper = workbook.getCreationHelper();
        final ClientAnchor anchor = helper.createClientAnchor();

        XSSFSheet sheet = workbook.getSheetAt(target.getInt(0));
        final Drawing drawing = sheet.createDrawingPatriarch();

        anchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
        anchor.setCol1(target.getJsonArray(1).getInt(0));
        anchor.setRow1(target.getJsonArray(1).getInt(1));
        anchor.setCol2(target.getJsonArray(2).getInt(0));
        anchor.setRow2(target.getJsonArray(2).getInt(1));
        
        final Picture pict = drawing.createPicture(anchor, pictureIndex);
        pict.resize();
    }

    private void compileTemplate(JsonArray cellData) {
        
        for (JsonValue chunkToInsert : cellData) {
            JsonArray target = ((JsonObject) chunkToInsert).getJsonArray("target");
            JsonArray data = ((JsonObject) chunkToInsert).getJsonArray("data");

            String chunkType;
            chunkType = (((JsonObject) chunkToInsert).containsKey("type") ?
                    ((JsonObject) chunkToInsert).getString("type") : "");

            switch (chunkType) {
                case "imageBase64":
                    addImageBase64ToSheet(target, data);
                    break;
                default:
                    populateTextSheet(target, data);  
            }
        }
    }

    private void recalculateAll() {
        XSSFFormulaEvaluator evaluator = XSSFFormulaEvaluator.create(workbook, IStabilityClassifier.TOTALLY_IMMUTABLE, UDFFinder.DEFAULT);
        evaluator.evaluateAll();
    }

    private void recalculateSheet(int index) {
        XSSFFormulaEvaluator evaluator = XSSFFormulaEvaluator.create(workbook, IStabilityClassifier.TOTALLY_IMMUTABLE, UDFFinder.DEFAULT);
        Sheet sheet = workbook.getSheetAt(index);
        String sheetName = sheet.getSheetName();
        System.out.println("Evaluating " + sheetName);
        for (Row r : sheet) {
            for (Cell c : r) {
                if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    evaluator.evaluateFormulaCell(c);
                }
            }
        }
        System.out.println("Done " + sheetName);
    }

    private boolean isCalcSheet(int index) {
        Sheet sheet = workbook.getSheetAt(index);
        String sheetName = sheet.getSheetName();
        return sheetName.contains("Calc");
    }

    private boolean isOutputSheet(int index) {
        Sheet sheet = workbook.getSheetAt(index);
        String sheetName = sheet.getSheetName();
        return sheetName.contains("Output");
    }

    private void streamWorkbook(OutputStream stream) throws IOException {
        workbook.write(stream);
        stream.close();
    }
    
    /**
     *
     * @param templatePath: URL of the template
     * @param JSONString: A JSON string Array with the new data. The structure is:
     *      [ { "type" : (String) Optional, if "imageBase64", will call a method to 
     *                   attach a image. Otherwise, will fill numbers
     *          "target" : 
     *              [
     *                  sheetNumber,
     *                  Upper Left coordinate [X, Y]
     *              ],
     *          "data": [[]] ==> Matrix in standard notation rows, columns
     *       }]
     * @throws IOException String templatePath, String JSONString
     */
    static public void compileAndStreamTemplate(String templatePath, String JSONString) throws IOException {
        JsonReader jsonReader = Json.createReader(new StringReader(JSONString));
        
        XlsxTemplate instance = new XlsxTemplate(templatePath);
        instance.compileTemplate(jsonReader.readArray());
        instance.recalculateAll();
        instance.streamWorkbook(System.out);
    }
    
    public static void main(String[] args) throws IOException {
        XlsxTemplate instance = new XlsxTemplate(args[0]);
        System.out.println("Loaded");
        XlsxTemplate.recalculateSAF(workbook);
        System.out.println("DONE");
        //instance.streamWorkbook(System.out);   
    }
     
    public static void recalculateSAF(XSSFWorkbook wb) {
        XSSFFormulaEvaluator evaluator = XSSFFormulaEvaluator.create(wb, IStabilityClassifier.TOTALLY_IMMUTABLE, UDFFinder.DEFAULT);
        for (Sheet sheet : wb) {
            System.out.println(sheet.getSheetName());
            IntStream rows = Arrays.stream(IntStream.range(0, sheet.getLastRowNum() + 1).toArray());
            rows.filter(rowNum -> sheet.getRow(rowNum) != null)
                .parallel()
                .forEach(rowNum -> {
                    Row r = sheet.getRow(rowNum);                    
                    for (Cell c : r) {
                        if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                            try {
                                evaluator.evaluateInCell(c);
                            } 
                            catch (Exception e) {
                              if (debug) {
                                  System.out.println("Failed at: " + sheet.getSheetName() + " r:" + rowNum + " c:" + c.getColumnIndex() + " F:" + c.getCellFormula());
                              }  
                            }
                        }
                    }
                });
        }
    }
}
