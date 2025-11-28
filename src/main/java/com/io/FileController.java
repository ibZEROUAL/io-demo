package com.io;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

@Controller
public class FileController {

    String fileName;

    @GetMapping
    public String input(){
        return "index";
    }

    @PostMapping("/uploadExcel")
    public ResponseEntity<FileSystemResource> uploadExcelFile(@RequestParam MultipartFile file) throws IOException {
        FileSystemResource resource;
        try(FileInputStream inputStream = new FileInputStream(toFile(file));
            BufferedWriter writer = new BufferedWriter(new FileWriter(fileName.split("\\.")[0]+".txt"));
            Workbook workbook = new XSSFWorkbook(inputStream)){

            Sheet sheet = workbook.getSheetAt(0);

            Map<Integer, List<String>> data = new HashMap<>();
            int i = 0;
            for (Row row : sheet) {
                data.put(i, new ArrayList<>());
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING: {
                            writer.write(cell.getStringCellValue()+"|");
                            break;
                        }
                        case NUMERIC: {
                            writer.write(Integer.valueOf((int) cell.getNumericCellValue())+"|");
                            break;
                        }
                        case BOOLEAN: {
                            writer.write(cell.getBooleanCellValue()+"|");
                            break;
                        }
                        case FORMULA: {
                            writer.write(cell.getCellFormula()+"|");
                            break;
                        }
                        default: data.get(i).add(" ");
                    }
                }
                writer.newLine();
                i++;
            }
            resource = new FileSystemResource(fileName.split("\\.")[0]+".txt");
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName.split("\\.")[0]+".txt" + "\"");
            headers.add(HttpHeaders.CACHE_CONTROL, "no-cache, no-store, must-revalidate");
            headers.add(HttpHeaders.PRAGMA, "no-cache");
            headers.add(HttpHeaders.EXPIRES, "0");

//            Path source = Path.of("./"+fileName);
//            Path target = Path.of("./input-files/"+fileName);
//            Files.move(source, target, StandardCopyOption.REPLACE_EXISTING);

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentType(MediaType.TEXT_PLAIN)
                    .body(resource);
        } catch (IOException e) {
            throw new IOException(e);
        }
    }

    private  File toFile(MultipartFile multipartFile) throws IOException {
        File convFile = new File(Objects.requireNonNull(multipartFile.getOriginalFilename()));
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(multipartFile.getBytes());
        }
        this.fileName = convFile.getName();
        return convFile;
    }

}
