/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package learnexcel;

/**
 *
 * @author Женя
 */
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
public class ExcelEditor {

    void createNewBook() throws IOException {
        Workbook MyWB = new XSSFWorkbook();
        Sheet MyFirstSheet = MyWB.createSheet("Первый лист");
        for (int i=0;i<=100;i++){
        Row MyFirstRow = MyFirstSheet.createRow(i);
        Cell CellHelloWorld = MyFirstRow.createCell(0);
        Cell HiRusCell = MyFirstRow.createCell(1);
        CellHelloWorld.setCellValue("Hello World!");
        HiRusCell.setCellValue("Здарова");
        }
        Path file_path = FileSystems.getDefault().getPath("FirstTry.xlsx");
        FileOutputStream stream = new FileOutputStream(new File(file_path.toString()));
        MyWB.write(stream);
        MyWB.close();
        
    }  
    void createOldBook() throws FileNotFoundException, IOException{
        Workbook OldWB = new HSSFWorkbook();
        Sheet FirstOldSheet = OldWB.createSheet("Первый старый лист");
        for (int i=0;i<=100;i++){
        Row MyFirstRow = FirstOldSheet.createRow(i);
        Cell FirstCell = MyFirstRow.createCell(0);
        FirstCell.setCellValue("Ярик");
        Cell Second = MyFirstRow.createCell(1);
        Second.setCellValue("Опять");
        Cell Third = MyFirstRow.createCell(2);
        Third.setCellValue("Проспал");
        }
        Path old_path = FileSystems.getDefault().getPath("Old.xls");
        FileOutputStream oldStream = new FileOutputStream(new File(old_path.toString()));
        try {
            OldWB.write(oldStream);
        } catch (IOException ex) {
            Logger.getLogger(ExcelEditor.class.getName()).log(Level.SEVERE, null, ex);
        }
        OldWB.close();
    }
    
}
