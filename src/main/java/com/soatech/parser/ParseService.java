package com.soatech.parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.web.multipart.MultipartFile;

import java.nio.file.Path;
import java.util.ArrayList;

public interface ParseService {
    Path createFile(MultipartFile file);

    void excelParser(Path path);

    void createExcel(Path path, int index, ArrayList<Cell> headers, ArrayList<Row> rowList);
}
