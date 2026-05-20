package com.io;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
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
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;

@Controller
public class FileController {
    private static final Pattern DIGITS_ONLY = Pattern.compile("\\D");

    @GetMapping
    public String input(){
        return "index";
    }

    @PostMapping("/uploadExcel")
    public ResponseEntity<FileSystemResource> uploadExcelFile(@RequestParam MultipartFile file) throws IOException {
        String originalName = Objects.requireNonNullElse(file.getOriginalFilename(), "output.xlsx");
        String baseName = originalName.contains(".")
                ? originalName.substring(0, originalName.lastIndexOf('.'))
                : originalName;
        String outputFileName = baseName + ".txt";

        String currentDateForName = getStringCurrentDate();
        String currentDate = getCurrentDate();
        String currentTime = getCurrentTime();

        List<String> lines = new ArrayList<>();
        String fileVersion = "0000";

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();

            for (Row row : sheet) {
                String line = toDelimitedLine(row, formatter);
                if (line.isEmpty()) {
                    continue;
                }

                lines.add(line);
                if ("0000".equals(fileVersion)) {
                    fileVersion = extractVersionFromRow(row, formatter);
                }
            }
        }

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputFileName))) {
            writer.write(
                    "@nom_fic:bvov7010020000" + currentDateForName + "00" + fileVersion + ".unl \n" +
                            "@des_fic : OV," + currentDate + "\n" +
                            "@dat_gen : " + currentDate + " \n" +
                            "@heur_gen : " + currentTime + " \n" +
                            "@cod_emet : 70 ASSOSICATION AHMED EL HANSALI                  \n" +
                            "@cod_dest : 1002 TR BENI MELLAL                                     \n" +
                            "@N_remise : 00" + fileVersion + " \n" +
                            "@nbr_enr : " + lines.size() + " \n" +
                            "@taille (octets) :  \n" +
                            "@utilisateur :  \n" +
                            "@Tel : \n");

            for (String line : lines) {
                writer.write(line);
                writer.newLine();
            }
        }

        FileSystemResource resource = new FileSystemResource(outputFileName);
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + outputFileName + "\"");
        headers.add(HttpHeaders.CACHE_CONTROL, "no-cache, no-store, must-revalidate");
        headers.add(HttpHeaders.PRAGMA, "no-cache");
        headers.add(HttpHeaders.EXPIRES, "0");

        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.TEXT_PLAIN)
                .body(resource);
    }

    private String toDelimitedLine(Row row, DataFormatter formatter) {
        StringBuilder line = new StringBuilder();
        boolean hasData = false;
        int lastCellIndex = row.getLastCellNum() - 1;
        int penultimateCellIndex = row.getLastCellNum() - 2;

        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            String value = cell == null ? "" : formatter.formatCellValue(cell).trim();
            if (cellIndex == lastCellIndex || cellIndex == penultimateCellIndex) {
                value = value.replace(" ", "");
            }

            if (!value.isEmpty()) {
                hasData = true;
            }
            line.append(value).append("|");
        }

        return hasData ? line.toString() : "";
    }

    private String extractVersionFromRow(Row row, DataFormatter formatter) {
        Cell versionCell = row.getCell(3);
        if (versionCell == null) {
            return "0000";
        }

        String digits = DIGITS_ONLY.matcher(formatter.formatCellValue(versionCell)).replaceAll("");
        if (digits.isEmpty()) {
            return "0000";
        }

        int numericVersion = Integer.parseInt(digits);
        return String.valueOf(numericVersion);
    }


    private String getStringCurrentDate(){
        LocalDate date = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMyy");
        return date.format(formatter);
    }

    private String getCurrentDate(){
        LocalDate date = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        return date.format(formatter);
    }

    private String getCurrentTime(){
        LocalTime time = LocalTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("HH:mm:ss");
        return time.format(formatter);
    }



}
