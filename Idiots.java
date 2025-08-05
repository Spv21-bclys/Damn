import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import com.fasterxml.jackson.databind.*;
import org.apache.hc.client5.http.classic.methods.*;
import org.apache.hc.client5.http.impl.classic.*;
import org.apache.hc.core5.http.io.entity.*;
import org.apache.hc.core5.http.*;

import java.io.*;
import java.util.*;

public class ExcelApiAutomation {

    private static final String INPUT_FILE = "input.xlsx";
    private static final String OUTPUT_FILE = "output.xlsx";

    private static final String TOKEN_URL = "https://your-tiaa-token-api.com/oauth/token";
    private static final String MAIN_API_URL = "https://your-main-api.com/resources";

    private static final String USERNAME = "your_username";
    private static final String PASSWORD = "your_password";
    private static final String CLIENT_ID = "your_client_id";

    private static final ObjectMapper mapper = new ObjectMapper();
    private static String accessToken = null;

    public static void main(String[] args) {
        try {
            List<Map<String, String>> results = new ArrayList<>();

            // Step 1: Read Excel
            try (FileInputStream fis = new FileInputStream(INPUT_FILE);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0);
                int leNameCol = -1;

                // Find the "le_name" column index
                Row headerRow = sheet.getRow(0);
                for (Cell cell : headerRow) {
                    if ("le_name".equalsIgnoreCase(cell.getStringCellValue())) {
                        leNameCol = cell.getColumnIndex();
                        break;
                    }
                }

                if (leNameCol == -1) throw new RuntimeException("le_name column not found!");

                // Step 2: Loop rows
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String leName = row.getCell(leNameCol).getStringCellValue();
                    System.out.println("Processing: " + leName);

                    // Fetch API response
                    Map<String, String> apiResult = callMainApi(leName);
                    apiResult.put("le_name", leName);

                    results.add(apiResult);
                }
            }

            // Step 3: Write output to Excel
            writeResultsToExcel(results);
            System.out.println("Process completed! Results saved to " + OUTPUT_FILE);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ---------------- TOKEN FUNCTION ----------------
    private static String getToken() throws IOException {
        if (accessToken != null) return accessToken;  // reuse token if already fetched

        try (CloseableHttpClient client = HttpClients.createDefault()) {
            HttpPost post = new HttpPost(TOKEN_URL);
            post.setHeader("Content-Type", "application/x-www-form-urlencoded");
            post.setHeader("Cookie", "your_cookie_here");

            String body = "username=" + USERNAME +
                          "&password=" + PASSWORD +
                          "&grant_type=password" +
                          "&client_id=" + CLIENT_ID;

            post.setEntity(new StringEntity(body));

            try (CloseableHttpResponse response = client.execute(post)) {
                Map<String, Object> json = mapper.readValue(response.getEntity().getContent(), Map.class);
                accessToken = (String) json.get("access_token");
                return accessToken;
            }
        }
    }

    // ---------------- MAIN API FUNCTION ----------------
    private static Map<String, String> callMainApi(String searchString) {
        Map<String, String> result = new HashMap<>();
        try (CloseableHttpClient client = HttpClients.createDefault()) {
            String token = getToken();

            HttpGet get = new HttpGet(MAIN_API_URL + "?searchString=" + searchString +
                    "&resources=your_resources&limit=10&include=your_include_param");

            get.setHeader("Authorization", "Bearer " + token);
            get.setHeader("tiaa-correlation-id", "your_correlation_id");
            get.setHeader("ip-user-context", "your_user_context");

            try (CloseableHttpResponse response = client.execute(get)) {
                if (response.getCode() == 401) { // token expired
                    accessToken = null;
                    return callMainApi(searchString); // retry once
                }

                Map<String, Object> json = mapper.readValue(response.getEntity().getContent(), Map.class);

                // Extract required fields (adjust according to your API response structure)
                result.put("csid", json.getOrDefault("csid", "").toString());
                result.put("countryOfIncorporation", json.getOrDefault("countryOfIncorporation", "").toString());
            }
        } catch (Exception e) {
            result.put("csid", "");
            result.put("countryOfIncorporation", "");
            result.put("error", e.getMessage());
        }
        return result;
    }

    // ---------------- WRITE OUTPUT EXCEL ----------------
    private static void writeResultsToExcel(List<Map<String, String>> results) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Results");

        // Create header
        Row header = sheet.createRow(0);
        String[] headers = {"le_name", "csid", "countryOfIncorporation", "error"};
        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }

        // Write data
        int rowIndex = 1;
        for (Map<String, String> row : results) {
            Row excelRow = sheet.createRow(rowIndex++);
            excelRow.createCell(0).setCellValue(row.getOrDefault("le_name", ""));
            excelRow.createCell(1).setCellValue(row.getOrDefault("csid", ""));
            excelRow.createCell(2).setCellValue(row.getOrDefault("countryOfIncorporation", ""));
            excelRow.createCell(3).setCellValue(row.getOrDefault("error", ""));
        }

        try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILE)) {
            workbook.write(fos);
        }
        workbook.close();
    }
}
