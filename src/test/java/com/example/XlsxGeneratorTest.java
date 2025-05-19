import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.json.JSONTokener;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

public class XlsxGeneratorTest {
    // Цвета для стилей
    private static final IndexedColors[] COLORS = {
            IndexedColors.GREY_25_PERCENT,    // серый (верхняя строка)
            IndexedColors.GREY_40_PERCENT,    // светло-серый (левый столбец)
            IndexedColors.PALE_BLUE,          // бледно-голубой (четные числовые)
            IndexedColors.LIGHT_BLUE,         // бледно-синий (нечетные числовые)
            IndexedColors.LIGHT_GREEN,        // бледно-зеленый (четные текстовые)
            IndexedColors.LIGHT_ORANGE,       // бледно-оранжевый (нечетные текстовые)
            IndexedColors.LIGHT_YELLOW,       // бледно-фиолетовый (даты)
    };

    // Стили для разных типов ячеек
    private Map<String, CellStyle> styles;
    private Workbook workbook;
    private boolean debugMode;

    public static void main(String[] args) {
        String outputPath = "."; // По умолчанию текущая директория
        boolean debug = false;

        // Обработка аргументов командной строки
        for (String arg : args) {
            if (arg.equals("d") || arg.equals("--debug")) {
                debug = true;
            } else if (!arg.startsWith("-")) {
                outputPath = arg;
            }
        }

        XlsxGeneratorTest generator = new XlsxGeneratorTest(debug);
        try {
            generator.generate(outputPath);
        } catch (Exception e) {
            System.err.println("Ошибка при генерации файла: " + e.getMessage());
            if (debug) e.printStackTrace();
        }
    }

    public XlsxGeneratorTest(boolean debugMode) {
        this.debugMode = debugMode;
        this.workbook = new XSSFWorkbook();
        this.styles = new HashMap<>();
        createDefaultStyles();
    }

    /**
     * Основной метод генерации файла
     */
    public void generate(String outputPath) throws Exception {
        if (debugMode) {
            System.out.println("Начало генерации XLSX файла...");
            System.out.println("Чтение данных из xlsx_data.json");
        }

        // Чтение JSON файла
        JSONObject jsonData = readJsonData();

        if (debugMode) {
            System.out.println("Данные успешно прочитаны:");
            System.out.println(jsonData.toString(2));
        }

        // Создание листов в книге
        JSONArray sheets = jsonData.getJSONArray("sheets");
        for (int i = 0; i < sheets.length(); i++) {
            JSONObject sheetData = sheets.getJSONObject(i);
            createSheet(sheetData);
        }

        // Генерация имени файла
        String fileName = generateFileName();
        String fullPath = outputPath + File.separator + fileName;

        // Создание директории, если она не существует
        Files.createDirectories(Paths.get(outputPath));

        // Сохранение файла
        try (FileOutputStream outputStream = new FileOutputStream(fullPath)) {
            workbook.write(outputStream);
        }

        if (debugMode) {
            System.out.println("Файл успешно сохранен: " + fullPath);
        }
    }

    /**
     * Чтение JSON данных из файла
     */
    private JSONObject readJsonData() throws Exception {
        try (InputStream inputStream = new FileInputStream("xlsx_data.json")) {
            return new JSONObject(new JSONTokener(inputStream));
        }
    }

    /**
     * Создание листа в книге на основе JSON данных
     */
    private void createSheet(JSONObject sheetData) {
        String sheetName = sheetData.optString("name", "Sheet " + (workbook.getNumberOfSheets() + 1));
        Sheet sheet = workbook.createSheet(sheetName);

        if (debugMode) {
            System.out.println("Создание листа: " + sheetName);
        }

        // Настройка ширины столбцов
        JSONArray columnWidths = sheetData.optJSONArray("columnWidths");
        if (columnWidths != null) {
            for (int i = 0; i < columnWidths.length(); i++) {
                int width = columnWidths.getInt(i);
                sheet.setColumnWidth(i, width * 256); // В Excel ширина измеряется в 1/256 символа
            }
        }

        // Настройка высоты строк
        JSONArray rowHeights = sheetData.optJSONArray("rowHeights");
        
        // Заполнение данными
        JSONArray data = sheetData.getJSONArray("data");
        for (int i = 0; i < data.length(); i++) {
            Row row = sheet.createRow(i);
            
            // Установка высоты строки, если указана
            if (rowHeights != null && i < rowHeights.length()) {
                row.setHeightInPoints(rowHeights.getFloat(i));
            }

            JSONArray rowData = data.getJSONArray(i);
            for (int j = 0; j < rowData.length(); j++) {
                Cell cell = row.createCell(j);
                Object cellData = rowData.get(j);
                
                // Определение типа данных ячейки
                if (cellData instanceof String) {
                    String value = (String) cellData;
                    if (value.startsWith("=")) {
                        cell.setCellFormula(value.substring(1));
                    } else {
                        cell.setCellValue(value);
                    }
                } else if (cellData instanceof Number) {
                    cell.setCellValue(((Number) cellData).doubleValue());
                } else if (cellData instanceof Boolean) {
                    cell.setCellValue((Boolean) cellData);
                }

                // Применение стиля к ячейке
                applyCellStyle(cell, i, j, cellData);
            }
        }

        // Авторазмер столбцов, если не задана ширина
        if (columnWidths == null) {
            for (int i = 0; i < data.getJSONArray(0).length(); i++) {
                sheet.autoSizeColumn(i);
            }
        }
    }

    /**
     * Применение стиля к ячейке на основе ее позиции и содержимого
     */
    private void applyCellStyle(Cell cell, int rowIndex, int colIndex, Object cellData) {
        String styleKey = "default";
        
        // Верхняя строка
        if (rowIndex == 0) {
            styleKey = "header";
        } 
        // Левый столбец
        else if (colIndex == 0) {
            styleKey = "firstColumn";
        }
        // Остальные ячейки
        else {
            boolean isEvenCol = colIndex % 2 == 0;
            
            if (cellData instanceof Number) {
                styleKey = isEvenCol ? "evenNumber" : "oddNumber";
            } else if (cellData instanceof String) {
                String value = (String) cellData;
                if (value.matches("\\d{4}-\\d{2}-\\d{2}")) { // Простая проверка на дату
                    styleKey = "date";
                } else {
                    styleKey = isEvenCol ? "evenText" : "oddText";
                }
            }
        }
        
        cell.setCellStyle(styles.get(styleKey));
    }

    /**
     * Создание стандартных стилей
     */
    private void createDefaultStyles() {
        // Общий стиль для всех ячеек
        CellStyle defaultStyle = workbook.createCellStyle();
        applyCommonStyle(defaultStyle);
        styles.put("default", defaultStyle);

        // Стиль для заголовков
        CellStyle headerStyle = workbook.createCellStyle();
        applyCommonStyle(headerStyle);
        headerStyle.setFillForegroundColor(COLORS[0].getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        styles.put("header", headerStyle);

        // Стиль для первого столбца
        CellStyle firstColumnStyle = workbook.createCellStyle();
        applyCommonStyle(firstColumnStyle);
        firstColumnStyle.setFillForegroundColor(COLORS[1].getIndex());
        firstColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font firstColumnFont = workbook.createFont();
        firstColumnFont.setBold(true);
        firstColumnStyle.setFont(firstColumnFont);
        firstColumnStyle.setAlignment(HorizontalAlignment.CENTER);
        firstColumnStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        firstColumnStyle.setWrapText(true);
        styles.put("firstColumn", firstColumnStyle);

        // Стили для числовых ячеек
        CellStyle evenNumberStyle = createDataStyle(COLORS[2]);
        styles.put("evenNumber", evenNumberStyle);
        CellStyle oddNumberStyle = createDataStyle(COLORS[3]);
        styles.put("oddNumber", oddNumberStyle);

        // Стили для текстовых ячеек
        CellStyle evenTextStyle = createDataStyle(COLORS[4]);
        styles.put("evenText", evenTextStyle);
        CellStyle oddTextStyle = createDataStyle(COLORS[5]);
        styles.put("oddText", oddTextStyle);

        // Стиль для ячеек с датами
        CellStyle dateStyle = createDataStyle(COLORS[6]);
        dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
        styles.put("date", dateStyle);
    }

    /**
     * Применение общих стилей (границы, отступы)
     */
    private void applyCommonStyle(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setIndention((short) 1);
    }

    /**
     * Создание стиля для данных с указанным цветом
     */
    private CellStyle createDataStyle(IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        applyCommonStyle(style);
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    /**
     * Генерация уникального имени файла
     */
    private String generateFileName() {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String randomId = UUID.randomUUID().toString().substring(0, 4);
        return "generated_" + timestamp + "_" + randomId + ".xlsx";
    }
}