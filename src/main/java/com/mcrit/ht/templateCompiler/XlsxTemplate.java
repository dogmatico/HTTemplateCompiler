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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;

import java.util.Arrays;
import java.util.stream.IntStream;
import java.util.Base64;
import java.util.HashMap;
import java.util.Iterator;

import javax.json.Json;
import javax.json.JsonArray;
import javax.json.JsonObject;
import javax.json.JsonReader;
import javax.json.JsonValue;
import org.apache.batik.transcoder.TranscoderException;

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

import org.apache.batik.transcoder.image.PNGTranscoder;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

/**
 * Class used in the HIGH-TOOL projecto to render XLSX templates.
 * It receives data from stdio as JSON and pipes the compiled template to 
 * stdout. Those streams are used to comunicate with the hosting environment: 
 * a Node.js child process.
 * @author Cristian Lorenzo i Martínez <cristian.lorenzo.martinez@gmail.com>
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
    
    private XSSFSheet getSheetByNameOrIndex(JsonArray target) {
        XSSFSheet sheet;
        
        switch (target.get(0).getValueType()) {
            case STRING:
                sheet = workbook.getSheet(target.getString(0));
                if (sheet == null) {
                    throw new java.lang.IllegalArgumentException("Sheet with name " + target.getString(0) + " not found.");
                }
                break;
            case NUMBER:
                sheet = workbook.getSheetAt(target.getInt(0));
                break;
            default:
                throw new RuntimeException("Invalid type in target(0). Must be a integer or a string.");
        }
                
        return sheet;
    }
    
    private void applyStyleToSheet(JsonArray target, JsonArray data) {
        XSSFSheet sheet = getSheetByNameOrIndex(target);
        
        data.forEach(it -> {
            JsonObject styleToApply = (JsonObject) it;
            XSSFCellStyle style = stylesDict.get(styleToApply.getString("style"));
            
            if (style == null) {
                throw new IllegalArgumentException("The style " + styleToApply.getString("style") + " does not exists. Unable to apply it to a sheet.");
            }
            
            JsonArray targets = styleToApply.getJsonArray("targets");
            
            targets.forEach(targetCells -> {
                int[] upperLeft = {
                    ((JsonArray) targetCells).getJsonArray(0).getInt(0),
                    ((JsonArray) targetCells).getJsonArray(0).getInt(1)
                };
                
                int[] lowerRight = {
                    ((JsonArray) targetCells).getJsonArray(1).getInt(0),
                    ((JsonArray) targetCells).getJsonArray(1).getInt(1)
                };
                
                for (int i = upperLeft[1]; i <= lowerRight[1]; i+= 1) {
                    XSSFRow row = sheet.getRow(i); 
                    if (row == null) {
                        sheet.createRow(i);
                        row = sheet.getRow(i);
                    }
                    for (int j = upperLeft[0]; j <= lowerRight[0]; j+= 1) {
                        XSSFCell cell = row.getCell(j, XSSFRow.CREATE_NULL_AS_BLANK);
                        cell.setCellStyle(style);
                    }
                }
            });
        });
    }
    
    
    private void populateTextSheet(JsonArray target, JsonArray data, String selectedStyle) {
        XSSFSheet sheet = getSheetByNameOrIndex(target);
        
        XSSFCellStyle style = stylesDict.get(selectedStyle);

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
                
                if (style != null) {
                    cell.setCellStyle(style);
                }
            }
        } 
    }
    
    private final HashMap<String, XSSFCellStyle> stylesDict = new HashMap();
    
    /**
     * 
     * @param JsonObject. The valid keys are:
     * fontSize : int,
     * fontColor : JsonArray[int, int, int] with rgb values,
     * fontBold: boolean,
     * fontItalic : boolean,
     * fontUnderline : string, valid values DOUBLE, DOUBLE_ACCOUNTING, NONE, SINGLE, SINGLE_ACCOUNTING,
     * backgroundColor : JsonArray[int, int, int] with rgb values,
     * border : {"color": array of JsonArray[int, int, int] with rgb values as size, 
     * "size" : int[] with [up, right, down, left] or [upDown, leftRight] or [all]}
     * "style" : string[], as others. Valid values DASH_DOT, DASH_DOT_DOT, DASHED, DOTTED, DOUBLE,
     * HAIR, MEDIUM, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, MEDIUM_DASHED, NONE, SLANTED_DASH_DOT, THICK, THIN;
     * "align" : String, valid values: CENTER, CENTER_SELECTION, DISTRIBUTED, FILL,
     * GENERAL, JUSTIFY, LEFT, RIGHT,
     * 
     */;
    
    /**
     * 
     * @param JsonObject. The valid keys are:
     * fontSize : int,
     * fontColor : JsonArray[int, int, int] with rgb values,
     * fontBold: boolean,
     * fontItalic : boolean,
     * fontUnderline : string, valid values DOUBLE, DOUBLE_ACCOUNTING, NONE, SINGLE, SINGLE_ACCOUNTING,
     * backgroundColor : JsonArray[int, int, int] with rgb values,
     * border : {"color": array of JsonArray[int, int, int] with rgb values as size, 
     * "size" : int[] with [up, right, down, left] or [upDown, leftRight] or [all]}
     * "style" : string[], as others. Valid values DASH_DOT, DASH_DOT_DOT, DASHED, DOTTED, DOUBLE,
     * HAIR, MEDIUM, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, MEDIUM_DASHED, NONE, SLANTED_DASH_DOT, THICK, THIN;
     * "align" : String, valid values: CENTER, CENTER_SELECTION, DISTRIBUTED, FILL,
     * GENERAL, JUSTIFY, LEFT, RIGHT,
     * 
     */
    private void parseStylesObject(JsonObject styles) {
        for (String styleName : styles.keySet()) {
            JsonObject style = styles.getJsonObject(styleName);
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            XSSFFont font = workbook.createFont();
            
            for (String styleProperty : style.keySet()) {
                // Create color
                XSSFColor color = null;

                switch (styleProperty) {
                    case "fontColor":
                    case "backgroundColor":
                        JsonArray rgb = style.getJsonArray(styleProperty);
                        color = new XSSFColor(new java.awt.Color(
                            rgb.getInt(0),
                            rgb.getInt(1),
                            rgb.getInt(2))
                        );
                        break;
                }

                switch (styleProperty) {
                    case "fontSize":
                        short ppt = (short) style.getInt(styleProperty);
                        font.setFontHeightInPoints(ppt);
                        break;
                    case "fontColor":
                        font.setColor(color);
                        break;
                    case "fontBold":
                        boolean setBold = style.getBoolean(styleProperty);
                        font.setBold(setBold);
                        if (setBold) {
                            font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                        } else {
                           font.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL); 
                        }

                        break;
                    case "fontItalic":
                        font.setItalic(style.getBoolean(styleProperty));
                        break;
                    case "fontUnderline":
                        font.setUnderline(FontUnderline.valueOf(style.getString(styleProperty)));
                        break;
                    case "backgroundColor":
                        cellStyle.setFillBackgroundColor(color);
                        cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                        break;
                    case "border":
                        JsonObject borderProperties = style.getJsonObject(styleProperty);
                        for (String borderProperty : borderProperties.keySet()) {
                            JsonArray borderStyleProp;
                            borderStyleProp = borderProperties.getJsonArray(borderProperty);

                            switch (borderProperty) {
                                case "color":
                                    XSSFColor[] colors = new XSSFColor[borderStyleProp.size()];

                                    for (int i = 0; i < borderStyleProp.size(); i += 1) {
                                        colors[i] = new XSSFColor(new java.awt.Color(
                                                    borderStyleProp.getJsonArray(i).getInt(0),
                                                    borderStyleProp.getJsonArray(i).getInt(1),
                                                    borderStyleProp.getJsonArray(i).getInt(2))
                                                );
                                    }

                                    if (colors != null) {
                                        switch (colors.length) {
                                            case 4:
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, colors[0]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, colors[1]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, colors[2]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, colors[3]);
                                                break;
                                            case 2:
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, colors[0]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, colors[0]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, colors[1]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, colors[1]);
                                                break;
                                            default:
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, colors[0]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, colors[0]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, colors[0]);
                                                cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, colors[0]);   
                                        }
                                    }
                                    break;
                                case "style" :
                                    switch (borderStyleProp.size()) {
                                        case 4:
                                            cellStyle.setBorderTop(BorderStyle.valueOf(borderStyleProp.getString(0)));
                                            cellStyle.setBorderRight(BorderStyle.valueOf(borderStyleProp.getString(1)));
                                            cellStyle.setBorderBottom(BorderStyle.valueOf(borderStyleProp.getString(2)));
                                            cellStyle.setBorderLeft(BorderStyle.valueOf(borderStyleProp.getString(3)));
                                            break;
                                        case 2:
                                            cellStyle.setBorderTop(BorderStyle.valueOf(borderStyleProp.getString(0)));
                                            cellStyle.setBorderBottom(BorderStyle.valueOf(borderStyleProp.getString(0)));
                                            cellStyle.setBorderRight(BorderStyle.valueOf(borderStyleProp.getString(1)));
                                            cellStyle.setBorderLeft(BorderStyle.valueOf(borderStyleProp.getString(1)));
                                            break;
                                        default:
                                            cellStyle.setBorderTop(BorderStyle.valueOf(borderStyleProp.getString(0)));
                                            cellStyle.setBorderRight(BorderStyle.valueOf(borderStyleProp.getString(1)));
                                            cellStyle.setBorderBottom(BorderStyle.valueOf(borderStyleProp.getString(2)));
                                            cellStyle.setBorderLeft(BorderStyle.valueOf(borderStyleProp.getString(3)));   
                                    }
                                    break;
                                default :
                                    throw new IllegalArgumentException("Unkown border property.");
                            }
                        }
                        break;
                    case "align":
                        cellStyle.setAlignment(HorizontalAlignment.valueOf(style.getString(styleProperty)));
                        break;
                    default : 
                        throw new IllegalArgumentException("Unkown property.");
                }
            }
            
            cellStyle.setFont(font);
            stylesDict.put(styleName, cellStyle);
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
    private void addImageBase64ToSheet(JsonArray target, JsonArray data) throws TranscoderException, IOException {
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
            case "svg":
            case "png":
               imageType = Workbook.PICTURE_TYPE_PNG;
               break;
            case "wmf":
               imageType = Workbook.PICTURE_TYPE_WMF;
               break;
           default:
               throw new RuntimeException("The provided format type is not supported.");
        }

        final byte[] decodedImg;
        if ("svg".equals(data.getString(0).toLowerCase())) {
            // Use batik to transform svg to png;
            PNGTranscoder transcoder = new PNGTranscoder();

            TranscoderInput input = new TranscoderInput(new ByteArrayInputStream(Base64.getDecoder().decode(data.getString(1))));
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            TranscoderOutput output = new TranscoderOutput(bos);

            transcoder.transcode(input, output);
            bos.flush();
            bos.close();
            decodedImg = ((ByteArrayOutputStream) output.getOutputStream()).toByteArray();

        } else {
            decodedImg = Base64.getDecoder().decode(data.getString(1));
        }

        final int pictureIndex = workbook.addPicture(decodedImg, imageType);

        final CreationHelper helper = workbook.getCreationHelper();
        final ClientAnchor anchor = helper.createClientAnchor();

        XSSFSheet sheet = getSheetByNameOrIndex(target);
        final Drawing drawing = sheet.createDrawingPatriarch();

        anchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
        anchor.setCol1(target.getJsonArray(1).getInt(0));
        anchor.setRow1(target.getJsonArray(1).getInt(1));
        anchor.setCol2(target.getJsonArray(2).getInt(0));
        anchor.setRow2(target.getJsonArray(2).getInt(1));

        final Picture pict = drawing.createPicture(anchor, pictureIndex);
        pict.resize();
    }

    private void compileTemplate(JsonArray cellData) throws TranscoderException, IOException {
        
        for (JsonValue chunkToInsert : cellData) {
            JsonArray target = ((JsonObject) chunkToInsert).getJsonArray("target");
            JsonArray data = ((JsonObject) chunkToInsert).getJsonArray("data");

            String selectedStyle = (((JsonObject) chunkToInsert).containsKey("style") ?
                ((JsonObject) chunkToInsert).getString("style") : 
                null);
                    
            String chunkType;
            chunkType = (((JsonObject) chunkToInsert).containsKey("type") ?
                    ((JsonObject) chunkToInsert).getString("type") : "");

            switch (chunkType) {
                case "styles":
                    applyStyleToSheet(target, data);
                    break;
                case "imageBase64":
                    addImageBase64ToSheet(target, data);
                    break;
                default:
                    populateTextSheet(target, data, selectedStyle);  
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
     * @param JsonArrayData: A JSON string Array with the new data. The structure is:
     *      [ { "type" : (String) Optional, if "imageBase64", will call a method to 
     *                   attach a image. Otherwise, will fill numbers
     *          "target" : 
     *              [
     *                  sheetNumber,
     *                  Upper Left coordinate [X, Y]
     *              ],
     *          "data": [[]] ==> Matrix in standard notation rows, columns
     *       }]
     * @param JsonObjectStyles : A JSON string object. Each key is a style. The keys of 
     * the style define properties.
     * @throws IOException String templatePath, String JSONString
     * @throws org.apache.batik.transcoder.TranscoderException
     */
    static public void compileAndStreamTemplate(String templatePath, String JsonArrayData, String JsonObjectStyles) throws IOException, TranscoderException {
        JsonArray data = Json.createReader(new StringReader(JsonArrayData)).readArray();
        JsonObject styles = Json.createReader(new StringReader(JsonObjectStyles)).readObject();
        
        
        XlsxTemplate instance = new XlsxTemplate(templatePath);
        instance.parseStylesObject(styles);
        instance.compileTemplate(data);
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
