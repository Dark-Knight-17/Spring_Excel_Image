/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.appoctet.ExcelImage;

import java.io.FileOutputStream;
import java.io.InputStream;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



@RestController	// This means that this class is a Controller
@RequestMapping(path="/api") // This means URL's start with /api (after Application path)

/**
 *
 * @author M
 */
public class ExcelController {
    
    @GetMapping(path="/create")
	public @ResponseBody String createFile() {
            
                try { 
  
            XSSFWorkbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Avengers");
            Row row1 = sheet.createRow(0);
            row1.createCell(0).setCellValue("BAT-MAN");
            Row row2 = sheet.createRow(1);
            row2.createCell(0).setCellValue("IRON-MAN");
            Row row3 = sheet.createRow(2);
            row3.createCell(0).setCellValue("SUPER-MAN");
            Row row4 = sheet.createRow(3);
            row4.createCell(0).setCellValue("SPIDER-MAN");
            Row row5 = sheet.createRow(4);
            row5.createCell(0).setCellValue("THOR");
            
            
            InputStream inputStream1 = ExcelController.class.getResourceAsStream("batman.jpg");
            InputStream inputStream2 = ExcelController.class.getResourceAsStream("ironman.jpg");
            InputStream inputStream3 = ExcelController.class.getResourceAsStream("superman.jpg");
            InputStream inputStream4 = ExcelController.class.getResourceAsStream("spiderman.jpg");
            InputStream inputStream5 = ExcelController.class.getResourceAsStream("thor.jpg");
            
            
            byte[] inputImageBytes1 = IOUtils.toByteArray(inputStream1);
            byte[] inputImageBytes2 = IOUtils.toByteArray(inputStream2);
            byte[] inputImageBytes3 = IOUtils.toByteArray(inputStream3);
            byte[] inputImageBytes4 = IOUtils.toByteArray(inputStream4);
            byte[] inputImageBytes5 = IOUtils.toByteArray(inputStream5);
            
            
            int inputImagePictureID1 = workbook.addPicture(inputImageBytes1, XSSFWorkbook.PICTURE_TYPE_JPEG);
            int inputImagePictureID2 = workbook.addPicture(inputImageBytes2, XSSFWorkbook.PICTURE_TYPE_JPEG);
            int inputImagePictureID3 = workbook.addPicture(inputImageBytes3, XSSFWorkbook.PICTURE_TYPE_JPEG);
            int inputImagePictureID4 = workbook.addPicture(inputImageBytes4, XSSFWorkbook.PICTURE_TYPE_JPEG);
            int inputImagePictureID5 = workbook.addPicture(inputImageBytes5, XSSFWorkbook.PICTURE_TYPE_JPEG);
            
            
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            
            XSSFClientAnchor ironManAnchor = new XSSFClientAnchor();
            XSSFClientAnchor spiderManAnchor = new XSSFClientAnchor();
            XSSFClientAnchor batManAnchor = new XSSFClientAnchor();
            XSSFClientAnchor superManAnchor = new XSSFClientAnchor();
            XSSFClientAnchor thorAnchor = new XSSFClientAnchor();
         
            
            ironManAnchor.setCol1(1); // Sets the column (0 based) of the first cell.
            ironManAnchor.setCol2(2); // Sets the column (0 based) of the Second cell.
            ironManAnchor.setRow1(0); // Sets the row (0 based) of the first cell.
            ironManAnchor.setRow2(1); // Sets the row (0 based) of the Second cell.
            
            spiderManAnchor.setCol1(1);
            spiderManAnchor.setCol2(2);
            spiderManAnchor.setRow1(1);
            spiderManAnchor.setRow2(2);
            
            batManAnchor.setCol1(1);
            batManAnchor.setCol2(2);
            batManAnchor.setRow1(2);
            batManAnchor.setRow2(3);
            
            superManAnchor.setCol1(1);
            superManAnchor.setCol2(2);
            superManAnchor.setRow1(3);
            superManAnchor.setRow2(4);
            
            thorAnchor.setCol1(1);
            thorAnchor.setCol2(2);
            thorAnchor.setRow1(4);
            thorAnchor.setRow2(5);
//            
            drawing.createPicture(ironManAnchor, inputImagePictureID1);
            drawing.createPicture(spiderManAnchor, inputImagePictureID2);
            drawing.createPicture(batManAnchor, inputImagePictureID3);
            drawing.createPicture(superManAnchor, inputImagePictureID4);
            drawing.createPicture(thorAnchor, inputImagePictureID5);
            
            
//            for (int i = 0; i < 5; i++) {
//                sheet.autoSizeColumn(i);
//            }
            
                FileOutputStream saveExcel = new FileOutputStream("test.xlsx");
                workbook.write(saveExcel);
                saveExcel.close();
                
                String msg="Excel File Created Succesfully";
                
                return msg;
                }
                catch(Exception e){
                
                return ("Sorry cannot create File ! Read the Error! "+e.toString());
                
                }
	 
	}
    
}
