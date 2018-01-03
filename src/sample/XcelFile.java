package sample;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class XcelFile {

    public static Workbook getBook(String filePath)
    {
        Workbook workbook;
        try
        {
            String extension = FilenameUtils.getExtension(filePath);
            FileInputStream excelFile = new FileInputStream(new File(filePath));

            System.out.println("got the extension: " + extension);

            if (extension.trim().equals("xls"))
            {
                System.out.println("start with xlS");
                workbook = new HSSFWorkbook(excelFile);
            }
            else if (extension.trim().equals("xlsx"))
            {
                System.out.println("start with xlSX");
                workbook = new XSSFWorkbook(excelFile);

            }
            else
            {
                System.out.println("fuck!");
                return null;
            }

        } catch (Exception ex)
        {
            workbook = null;
            ex.printStackTrace();
        }

        return workbook;
    }

    public static void main(String[] args) {
        try {
            Workbook workbook = new XSSFWorkbook(new FileInputStream(new File("/Users/luis/Desktop/teee.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
