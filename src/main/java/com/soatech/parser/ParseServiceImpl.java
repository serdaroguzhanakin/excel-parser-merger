package com.soatech.parser;

import com.soatech.helper.ExcelHelper;
import com.soatech.helper.IOHelper;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;

@Service
public class ParseServiceImpl implements ParseService {
    private static String PARSED_FOLDER = System.getProperty("user.dir") + "/parsed/";
    private static int ROW_COUNT = 100;


    @Override
    public Path createFile(MultipartFile file) {
        try {
            byte[] bytes = file.getBytes();
            Path path = Paths.get(PARSED_FOLDER + file.getOriginalFilename());
            Files.write(path, bytes);
            return path;
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }
    }

    @Override
    public void excelParser(Path path) {
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path.toString()));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;

            //get headers
            row = sheet.getRow(0);

            ArrayList<Cell> headers = new ArrayList<>();
            for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                headers.add(row.getCell(i));
            }

            ArrayList<Row> bodyRows = new ArrayList<>();
            int index = 0;
            boolean isLastLine = false;
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowCount; i++) {
                if (sheet.getRow(i).getCell(0) == null) {
                    isLastLine = true;
                }
                if (!isLastLine) {
                    bodyRows.add(sheet.getRow(i));
                }
                if (i % ROW_COUNT == 0 || i == rowCount - 1 || isLastLine) {
                    createExcel(path, index, headers, bodyRows);
                    bodyRows.clear();
                    index++;
                }
                if (isLastLine) {
                    break;
                }
            }

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    @Override
    public void createExcel(Path path, int index, ArrayList<Cell> headers, ArrayList<Row> bodyRows) {
        if (ExcelHelper.isRowListEmpty(bodyRows)) return;

        Workbook workbook = ExcelHelper.createWorkBook();

        Sheet sheet = ExcelHelper.createSheet(workbook, String.valueOf(index));

        ExcelHelper.createHeader(sheet, headers);

        ExcelHelper.createBody(sheet, bodyRows, true);

        ExcelHelper.resizeAllColumns(sheet, headers);

        IOHelper.saveFile(createParsedFolderDestinationPath(index, path), workbook);
    }

    private Path createParsedFolderDestinationPath(int index, Path path) {
        return Paths.get(
                PARSED_FOLDER +
                        index +
                        "_" +
                        path.getFileName()
        );
    }

}
