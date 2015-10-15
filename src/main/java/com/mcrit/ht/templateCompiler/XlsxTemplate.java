/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mcrit.ht.templateCompiler;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;
import javax.json.Json;
import javax.json.JsonArray;
import javax.json.JsonObject;
import javax.json.JsonReader;
import javax.json.JsonValue;
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
    
    public XlsxTemplate (String fileName) throws FileNotFoundException, IOException {
        FileInputStream fis = new FileInputStream(fileName);
        try {
	    workbook = new XSSFWorkbook(fis);
	} finally {
	    fis.close();
	}
    }
    
    private void compileTemplate(JsonArray cellData) {
        
        for (JsonValue chunkToInsert : cellData) {
            JsonArray target = ((JsonObject) chunkToInsert).getJsonArray("target");
            JsonArray data = ((JsonObject) chunkToInsert).getJsonArray("data");
            
            XSSFSheet sheet = workbook.getSheetAt(target.getInt(0));
            
            for (int j = 0, ln2 = data.size(); j < ln2; j++) {
                XSSFRow row = sheet.getRow(j + target.getJsonArray(1).getInt(1));
                if (row == null) {
                    sheet.createRow(j + target.getJsonArray(1).getInt(1));
                    row = sheet.getRow(j + target.getJsonArray(1).getInt(1));
                }
                for (int k = 0, ln3 = data.getJsonArray(j).size(); k < ln3; k++) {
                    XSSFCell cell = row.getCell(k + target.getJsonArray(1).getInt(0), XSSFRow.CREATE_NULL_AS_BLANK);
                    try {
                        cell.setCellValue(data.getJsonArray(j).getJsonNumber(k).doubleValue());
                    }
                    catch (ClassCastException ex) {
                        cell.setCellValue(data.getJsonArray(j).getString(k));
                    }
                }
            } 
        }
    }

    private void recalculateAll() {
        XSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
    }
    
    private void streamWorkbook(OutputStream stream) throws IOException {
        workbook.write(stream);
        stream.close();
    }
        
    /**
     *
     * @param templatePath: URL of the template
     * @param JSONString: A JSON string Array with the new data. The structure is:
     *      [ {target : [
     *          sheetNumber,
     *          Upper Left coordinate [X, Y]
     *        ],
     *        data: [[]] ==> Matrix in standard notation rows, columns
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
}
