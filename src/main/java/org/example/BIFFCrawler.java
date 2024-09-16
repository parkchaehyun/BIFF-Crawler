import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.FileOutputStream;
import java.io.IOException;

public class BIFFCrawler {
    public static void main(String[] args) {
        try {
            // Create a new Excel workbook
            Workbook workbook = new XSSFWorkbook();
            CreationHelper createHelper = workbook.getCreationHelper();

            // Create a new Excel sheet
            Sheet sheet = workbook.createSheet("Movie Schedule");

            for (int i = 0; i < 10; i++) {
                sheet.setColumnWidth(i, 256 * 26);
            }

            // Parse the HTML content
            int day = 6;
            Document doc = Jsoup.connect("https://www.biff.kr/kor/html/schedule/date.asp?day1=" + day).get();
            Elements theaterItems = doc.select(".sch_li");

            CellStyle style = workbook.createCellStyle();
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);

            int rowIdx = 0;

            // Create a header row
            Row headerRow = sheet.createRow(rowIdx++);
            headerRow.createCell(0).setCellValue("극장");
            headerRow.getCell(0).setCellStyle(style);
            for (int i = 1; i <= 6; i++) {
                headerRow.createCell(i).setCellValue(i + "회");
                headerRow.getCell(i).setCellStyle(style);
            }

            // Iterate through the theater items
            for (Element theaterItem : theaterItems) {
                Element theaterElement = theaterItem.selectFirst(".sch_li_tit");

                if (theaterElement != null) {
                    String theater = theaterElement.text();
                    System.out.println("parsing movies from " + theater);
                    // Create a new row and fill in the data
                    Row row = sheet.createRow(rowIdx++);
                    row.setHeightInPoints(73);
                    row.createCell(0).setCellValue(theater);

                    row.getCell(0).setCellStyle(style);

                    // Iterate through the movies within the theater
                    Elements movieItems = theaterItem.select(".sch_it");


                    for (Element movieItem : movieItems) {
                        Element timeElement = movieItem.selectFirst(".time.en");
                        Element codeElement = movieItem.selectFirst(".code.en");
                        Element titleElement = movieItem.selectFirst(".film_tit_kor");

                        if (timeElement != null && codeElement != null && titleElement != null) {
                            String time = timeElement.text();
                            String code = codeElement.text();
                            String title = titleElement.text();

                            String className = movieItem.className();
                            int itNumber = extractItNumber(className);

                            // Extract the URL of the individual movie page
                            Element movieLinkElement = titleElement.parent();
                            String moviePageUrl = movieLinkElement.attr("href");

                            // Parse the HTML content of the individual movie page
                            Document moviePageDoc = Jsoup.connect("https://www.biff.kr" + moviePageUrl).get();

                            // Locate and extract the runtime information
                            Element runtimeElement = moviePageDoc.selectFirst("li.en:contains(러닝타임)");
                            if (runtimeElement != null) {
                                String runtimeText = runtimeElement.text();
                                String numericRuntime = runtimeText.replaceAll("[^0-9]", "");
                                if (!numericRuntime.isEmpty()) {
                                    int runtimeValue = Integer.parseInt(numericRuntime);

                                    // Parse the start time (hh:mm)
                                    String[] timeParts = time.split(":");
                                    if (timeParts.length == 2) {
                                        int startHour = Integer.parseInt(timeParts[0]);
                                        int startMinute = Integer.parseInt(timeParts[1]);

                                        // Calculate the end time
                                        int endHour = startHour + (runtimeValue / 60);
                                        int endMinute = startMinute + (runtimeValue % 60);

                                        // Handle overflow if minutes exceed 60
                                        if (endMinute >= 60) {
                                            endHour += 1;
                                            endMinute %= 60;
                                        }

                                        // Format the start and end times into a new time range string
                                        String startTime = String.format("%02d:%02d", startHour, startMinute);
                                        String endTime = String.format("%02d:%02d", endHour, endMinute);
                                        time = startTime + "-" + endTime;

                                        // Now 'timeRange' contains the time range (e.g., "20:00-21:54")
                                    }
                                }
                            }

                            // Concatenate code, title, and time with newline character
                            StringBuilder moviesDetails = new StringBuilder();
                            moviesDetails.append(code).append("\n").append(title).append("\n").append(time).append("\n→ ");
                            row.createCell(itNumber).setCellValue(moviesDetails.toString());
                            row.getCell(itNumber).setCellStyle(style);
                        }
                    }

                }
            }

            // Write the Excel data to a file
            FileOutputStream fileOut = new FileOutputStream("movie_schedule_" + day + ".xlsx");
            workbook.write(fileOut);
            fileOut.close();

            // Close the workbook
            workbook.close();

            System.out.println("Excel file created successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static int extractItNumber(String className) {
        // Extract the number after 'sch_it' in the class name (e.g., "sch_it sch_it4")
        String[] parts = className.split("\\s+");
        for (String part : parts) {
            if (part.startsWith("sch_it")) {
                String numberPart = part.substring(6); // After 'sch_it'
                if (numberPart.matches("\\d+")) {
                    try {
                        return Integer.parseInt(numberPart);
                    } catch (NumberFormatException e) {
                        // Continue to next part if parsing fails
                        continue;
                    }
                }
            }
        }
        return 0; // Default value if not found
    }
}
