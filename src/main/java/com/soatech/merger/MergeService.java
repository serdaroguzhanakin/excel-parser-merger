package com.soatech.merger;

import org.springframework.web.multipart.MultipartFile;

import java.nio.file.Path;
import java.util.List;

public interface MergeService {
    Path createFile(MultipartFile file, int index);

    void appendFile(List<Path> paths);
}
