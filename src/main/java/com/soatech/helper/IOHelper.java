package com.soatech.helper;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;

public class IOHelper {

    public static void saveFile(Path destinationPath, Workbook workbook) {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(destinationPath.toString());
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }
    }
}
