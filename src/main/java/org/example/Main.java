package org.example;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.net.URI;
import java.net.http.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.*;
import java.time.LocalDateTime;
import java.util.Base64;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

public class Main {

    private static final ObjectMapper mapper = new ObjectMapper();
    private static final HttpClient client = HttpClient.newHttpClient();

    public static void main(String[] args) {

        ScheduledExecutorService scheduler =
                Executors.newSingleThreadScheduledExecutor();

        scheduler.scheduleAtFixedRate(() -> {
            try {
                System.out.println("\n==============================");
                System.out.println("Start Import: " + LocalDateTime.now());

                runImport();

                System.out.println("Import fertig");

            } catch (Exception e) {
                System.out.println("Fehler im Import:");
                e.printStackTrace();
            }
        }, 0, 1, TimeUnit.MINUTES);
    }

    private static void runImport() throws Exception {

       // Credentials einfügen !!!!
        String username = "USER";
        String password = "PASSWORD";

        if (username == null || password == null) {
            throw new RuntimeException("Missing env vars SKY_USER / SKY_PASS");
        }

        // 1. TOKEN holen

        String form = "grant_type=password"
                + "&username=" + username
                + "&password=" + password;

        HttpRequest tokenReq = HttpRequest.newBuilder()
                .uri(URI.create("https://admin.skyresponse.com/api/v2/token"))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(form))
                .build();

        HttpResponse<String> tokenRes = client.send(tokenReq, HttpResponse.BodyHandlers.ofString());

        String token = mapper.readTree(tokenRes.body())
                .get("access_token").asText();

        // 2. REPORT GENERIEREN

        String payload = """
        {
          "language":"de-DE-x-care-curena",
          "reportArguments":{
            "runMode":"BySearchCriteria",
            "exactMatch":false,
            "nameOnly":false,
            "entityIds":[],
            "organizationIds":[],
            "searchText":""
          },
          "outputFormat":"EXCEL",
          "reportId":9
        }
        """;

        HttpRequest genReq = HttpRequest.newBuilder()
                .uri(URI.create("https://admin.skyresponse.com/api/v2/reports/generate"))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/json")
                .POST(HttpRequest.BodyPublishers.ofString(payload))
                .build();

        HttpResponse<String> genRes = client.send(genReq, HttpResponse.BodyHandlers.ofString());

        String fileId = genRes.body().replace("\"", "").trim();

        // 3. DOWNLOAD POLLING

        String downloadUrl = "https://admin.skyresponse.com/api/v2/reports/download/"
                + fileId + "?language=de-DE-x-care-curena";

        JsonNode json = null;

        for (int i = 0; i < 15; i++) {

            HttpRequest downloadReq = HttpRequest.newBuilder()
                    .uri(URI.create(downloadUrl))
                    .header("Authorization", "Bearer " + token)
                    .GET()
                    .build();

            HttpResponse<String> res = client.send(downloadReq, HttpResponse.BodyHandlers.ofString());
            String body = res.body();

            if (!body.contains("FileNotFoundException")) {
                json = mapper.readTree(body);
                break;
            }

            System.out.println("Warte auf Datei... Versuch " + (i + 1));
            Thread.sleep(2000);
        }

        if (json == null) {
            System.out.println("Datei nicht verfügbar");
            return;
        }

        String fileName = json.get("fileName").asText();
        String base64 = json.get("fileData").asText();

        Files.write(Path.of(fileName), Base64.getDecoder().decode(base64));


        // 4. SQLITE

        Connection conn = DriverManager.getConnection("jdbc:sqlite:report.db");
        Statement stmt = conn.createStatement();

        stmt.execute("DROP TABLE IF EXISTS report");

        // 5. EXCEL LESEN

        FileInputStream fis = new FileInputStream(fileName);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        int headerRowIndex = 5;
        Row headerRow = sheet.getRow(headerRowIndex);

        int cols = headerRow.getLastCellNum();


        // 6. TABLE CREATE

        StringBuilder create = new StringBuilder("CREATE TABLE report (");

        for (int i = 0; i < cols; i++) {

            String raw = headerRow.getCell(i) != null
                    ? headerRow.getCell(i).toString()
                    : "col" + i;

            String col = safeColumnName(raw, i);

            create.append("\"").append(col).append("\" TEXT");

            if (i < cols - 1) create.append(", ");
        }

        create.append(");");

        stmt.execute(create.toString());


        // 7. INSERT

        StringBuilder insert = new StringBuilder("INSERT INTO report VALUES (");

        for (int i = 0; i < cols; i++) {
            insert.append("?");
            if (i < cols - 1) insert.append(", ");
        }

        insert.append(")");

        PreparedStatement ps = conn.prepareStatement(insert.toString());

        int rowIndex = 0;
        int inserted = 0;

        for (Row row : sheet) {

            rowIndex++;
            if (rowIndex <= 5) continue;
            if (row == null) continue;

            boolean empty = true;

            for (int i = 0; i < cols; i++) {

                Cell cell = row.getCell(i);
                String value = getCellValue(cell);

                if (value != null && !value.isBlank()) {
                    empty = false;
                }

                ps.setString(i + 1, value);
            }

            if (empty) continue;

            ps.addBatch();
            inserted++;

            if (inserted % 500 == 0) {
                ps.executeBatch();
            }
        }

        ps.executeBatch();

        workbook.close();
        conn.close();

        System.out.println("Importiert: " + inserted + " Zeilen");
    }

    private static String safeColumnName(String input, int index) {

        if (input == null) return "col" + index;

        String c = input.trim().toLowerCase()
                .replace(" ", "_")
                .replace("-", "_")
                .replace(".", "_")
                .replace("/", "_")
                .replace("\\", "_");

        c = c.replaceAll("[^a-z0-9_]", "");

        return c.isBlank() ? "col" + index : c;
    }

    private static String getCellValue(Cell cell) {

        if (cell == null) return null;

        try {
            return switch (cell.getCellType()) {

                case STRING -> cell.getStringCellValue();

                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        yield cell.getDateCellValue().toString();
                    }
                    yield String.valueOf(cell.getNumericCellValue());
                }

                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());

                case FORMULA -> {
                    try {
                        yield cell.getStringCellValue();
                    } catch (Exception e) {
                        yield String.valueOf(cell.getNumericCellValue());
                    }
                }

                default -> null;
            };
        } catch (Exception e) {
            return null;
        }
    }
}
