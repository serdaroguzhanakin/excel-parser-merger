package com.soatech.merger;

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
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

@Service
public class MergeServiceImpl implements MergeService {
    private static String MERGED_FOLDER = System.getProperty("user.dir") + "/merged/";

    private static DateTimeFormatter formatter =
            DateTimeFormatter.ofPattern("yyyy-MM-dd")
                    .withLocale(Locale.UK)
                    .withZone(ZoneId.systemDefault());

    private static String FILE_SUFFIX = ".xls";

    @Override
    public Path createFile(MultipartFile file, int index) {
        try {
            Path path = Paths.get(MERGED_FOLDER +
                    formatter.format(Instant.now()) +
                    "_" +
                    index +
                    FILE_SUFFIX);

            System.out.println("1-" + path.toString());

            Files.write(
                    path,
                    file.getBytes()
            );

            return path;
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }
    }

    @Override
    public void appendFile(List<Path> paths) {
        ArrayList<Cell> headers = new ArrayList<>();
        ArrayList<Row> bodyRows = new ArrayList<>();

        try {
            for (Path path : paths) {
                POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path.toString()));
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                HSSFSheet sheet = wb.getSheetAt(0);
                HSSFRow row;

                //get headers
                row = sheet.getRow(0);

                if (headers.isEmpty()) {
                    for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                        headers.add(row.getCell(i));
                    }
                }
                boolean isLastLine = false;
                int rowCount = sheet.getPhysicalNumberOfRows();
                for (int i = 1; i < rowCount; i++) {
                    if (sheet.getRow(i).getCell(0) == null) {
                        isLastLine = true;
                    }
                    if (!isLastLine) {
                        bodyRows.add(sheet.getRow(i));
                    }
                    if (isLastLine) {
                        break;
                    }
                }
            }

            createExcel(headers, bodyRows);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        }
    }

    private void createExcel(ArrayList<Cell> headers, ArrayList<Row> bodyRows) {
        if (ExcelHelper.isRowListEmpty(bodyRows)) return;

        Workbook workbook = ExcelHelper.createWorkBook();

        Sheet sheet = ExcelHelper.createSheet(workbook, "sheet-1");

        ExcelHelper.createHeader(sheet, headers);

        ExcelHelper.createBody(sheet, bodyRows, false);

        ExcelHelper.resizeAllColumns(sheet, headers);

        IOHelper.saveFile(createMergedFolderDestinationPath(), workbook);
    }

    private Path createMergedFolderDestinationPath() {
        return Paths.get(
                MERGED_FOLDER +
                        "LAST_" +
                        formatter.format(Instant.now()) +
                        FILE_SUFFIX
        );
    }
}
