/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mcrit.ht.templateCompiler;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import javax.json.JsonArray;
import javax.json.JsonObject;
import javax.json.JsonValue;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author Cristian Lorenzo <cristian.lorenzo.martinez@gmail.com>
 */
public class XlsxTemplate {
    private static HSSFWorkbook workbook;
    
    public XlsxTemplate (String fileName) throws FileNotFoundException, IOException {
        FileInputStream fis = new FileInputStream(fileName);
        try {
	    workbook = new HSSFWorkbook(fis);
	} finally {
	    fis.close();
	}
    }
    
    private void compileTemplate(JsonArray cellData) {
        
        for (JsonValue chunkToInsert : cellData) {
            JsonArray target = ((JsonObject) chunkToInsert).getJsonArray("target");
            JsonArray data = ((JsonObject) chunkToInsert).getJsonArray("data");
            
            HSSFSheet sheet = workbook.getSheetAt(target.getInt(0));
            
            for (int j = target.getJsonArray(1).getInt(1), ln2 = target.getJsonArray(1).getInt(1) + data.size(); j < ln2; j++) {
                HSSFRow row = sheet.getRow(j);
                if (row == null) {
                    sheet.createRow(j);
                }
                for (int k = target.getJsonArray(1).getInt(0), ln3 = target.getJsonArray(1).getInt(0) + data.getJsonArray(j).size(); k < ln2; k++) {
                    HSSFCell cell = row.getCell(k, HSSFRow.CREATE_NULL_AS_BLANK);
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
}
