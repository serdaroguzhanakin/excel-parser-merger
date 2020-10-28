package com.soatech.controller;

import com.soatech.merger.MergeService;
import com.soatech.parser.ParseService;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

@Controller
public class FileController {

    private ParseService parseService;
    private MergeService mergeService;

    public FileController(ParseService parseService,
                          MergeService mergeService) {
        this.parseService = parseService;
        this.mergeService = mergeService;
    }

    @GetMapping("/")
    public String index() {
        return "home";
    }

    @GetMapping("/home")
    public String home() {
        return "home";
    }

    @GetMapping("/parse")
    public String parse() {
        return "parse";
    }

    @GetMapping("/merge")
    public String merge() {
        return "merge";
    }

    @PostMapping("/parse")
    public String singleFileUpload(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) {
        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("message", "Please select a file to parse");
            redirectAttributes.addFlashAttribute("isSuccess", false);
            return "redirect:/status";
        }

        try {
            Path path = parseService.createFile(file);

            parseService.excelParser(path);

            redirectAttributes.addFlashAttribute("message",
                    "You successfully parsed '" + file.getOriginalFilename() + "'");
            redirectAttributes.addFlashAttribute("isSuccess", true);
        } catch (Exception e) {
            redirectAttributes.addFlashAttribute("message", e.getMessage());
            redirectAttributes.addFlashAttribute("isSuccess", false);
        }

        return "redirect:/status";
    }

    @PostMapping("/merge")
    public String mergeFiles(@RequestParam("files") List<MultipartFile> files, RedirectAttributes redirectAttributes) throws IOException {
        try {
            if (files.isEmpty()) {
                redirectAttributes.addFlashAttribute("message", "Please select files to merge");
                redirectAttributes.addFlashAttribute("isSuccess", false);
                return "redirect:/status";
            }

            MultipartFile firstFile = files.stream().findFirst().get();
            List<Path> paths = new ArrayList<>();
            int index = 0;
            for (MultipartFile file : files) {
                paths.add(mergeService.createFile(file, index++));
            }

            mergeService.appendFile(paths);

            redirectAttributes.addFlashAttribute("message", "You successfully merged files");
            redirectAttributes.addFlashAttribute("isSuccess", true);
        } catch (Exception e) {
            redirectAttributes.addFlashAttribute("message", e.getMessage());
            redirectAttributes.addFlashAttribute("isSuccess", false);
        }
        return "redirect:/status";
    }

    @GetMapping("/status")
    public String status() {
        return "status";
    }

}