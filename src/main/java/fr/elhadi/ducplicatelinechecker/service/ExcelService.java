package fr.elhadi.ducplicatelinechecker.service;

import fr.elhadi.ducplicatelinechecker.model.Element;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.EnumMap;
import java.util.stream.IntStream;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

public class ExcelService {

    private static final Logger logger = LogManager.getLogger();

    private enum Color {GREEN, YELLOW, RED}

    public static void generateExcelFile(List<Element> elementsANotPresentInB, List<Element> elementsBNotPresentInA) throws IOException {
        LocalDateTime localDateTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyy-HH.mm.ss");
        String fileName = "results-" + localDateTime.format(formatter) + ".xlsx";
        try (Workbook workbook = new XSSFWorkbook();
             OutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            Sheet sheet = prepareSheet(workbook);
            prepareHeader(workbook, sheet);
            EnumMap<Color, CellStyle> styles = createStyles(workbook);

            int index = 2;
            for (Element element : elementsANotPresentInB) {
                Element elementB = getSameItemsInBWithDifferentLeadtime(elementsBNotPresentInA, element).orElse(null);
                fillRow(sheet, styles, index, element, elementB);
                index++;
            }
            saveExcelFile(workbook, fileName);
        }
    }

    private static Sheet prepareSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("Items difference");
        IntStream.range(0, 28).forEach(i -> sheet.setColumnWidth(i, 4000));
        return sheet;
    }

    private static void prepareHeader(Workbook workbook, Sheet sheet) {
        Row header = sheet.createRow(0);
        CellStyle headerStyle = createHeaderStyle(workbook);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        fillCell(headerStyle, header, 0, "DppType");
        Cell headerCell;

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 3));
        headerCell = header.createCell(2);
        headerCell.setCellValue("DppThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 5));
        headerCell = header.createCell(4);
        headerCell.setCellValue("DppSubThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 7));
        headerCell = header.createCell(6);
        headerCell.setCellValue("CustomerType");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 8, 9));
        headerCell = header.createCell(8);
        headerCell.setCellValue("CustomerThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 10, 11));
        headerCell = header.createCell(10);
        headerCell.setCellValue("CustomerSubThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 12, 13));
        headerCell = header.createCell(12);
        headerCell.setCellValue("FinalType");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 14, 15));
        headerCell = header.createCell(14);
        headerCell.setCellValue("FinalThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 16, 17));
        headerCell = header.createCell(16);
        headerCell.setCellValue("FinalSubThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 18, 19));
        headerCell = header.createCell(18);
        headerCell.setCellValue("ItemNumber");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 20, 21));
        headerCell = header.createCell(20);
        headerCell.setCellValue("CurrentWeek");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 22, 23));
        headerCell = header.createCell(22);
        headerCell.setCellValue("CurrentYear");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 24, 25));
        headerCell = header.createCell(24);
        headerCell.setCellValue("replenishment");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 26, 27));
        headerCell = header.createCell(26);
        headerCell.setCellValue("provisions");
        headerCell.setCellStyle(headerStyle);

        Row subHeader = sheet.createRow(1);
        for (int i = 0; i < 28; i += 2) {
            headerCell = subHeader.createCell(i);
            headerCell.setCellValue("original");
            headerCell.setCellStyle(headerStyle);

            headerCell = subHeader.createCell(i + 1);
            headerCell.setCellValue("optimized");
            headerCell.setCellStyle(headerStyle);

        }
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);

        createHeaderFont((XSSFWorkbook) workbook, headerStyle);

        return headerStyle;
    }

    private static void createHeaderFont(XSSFWorkbook workbook, CellStyle headerStyle) {
        XSSFFont font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 12);
        headerStyle.setFont(font);
    }

    private static EnumMap<Color, CellStyle> createStyles(Workbook workbook) {
        EnumMap<Color, CellStyle> styles = new EnumMap<>(Color.class);
        styles.put(Color.GREEN, createStyle(workbook, IndexedColors.LIGHT_GREEN.getIndex()));
        styles.put(Color.YELLOW, createStyle(workbook, IndexedColors.LIGHT_YELLOW.getIndex()));
        styles.put(Color.RED, createStyle(workbook, IndexedColors.RED.getIndex()));
        return styles;
    }

    private static CellStyle createStyle(Workbook workbook, short color) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("@"));
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        return style;
    }

    private static Optional<Element> getSameItemsInBWithDifferentLeadtime(List<Element> elementsBNotPresentInA, Element element) {
        return elementsBNotPresentInA.stream().filter(b -> element.getDppType().equals(b.getDppType())
                && b.getDppThird().equals(element.getDppThird())
                && b.getDppSubThird().equals(element.getDppSubThird())
                && b.getCustomerType().equals(element.getCustomerType())
                && b.getCustomerThird().equals(element.getCustomerThird())
                && b.getCustomerSubTihrd().equals(element.getCustomerSubTihrd())
                && b.getFinalType().equals(element.getFinalType())
                && b.getFinalThird().equals(element.getFinalThird())
                && b.getFinalSubThird().equals(element.getFinalSubThird())
                && b.getItemNumber().equals(element.getItemNumber())).findFirst();
    }

    private static void fillRow(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB) {
        Row row = sheet.createRow(index);
        fillCells(sheet, styles, index, element, elementB, row);
    }

    private static void fillCells(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (elementB != null) {
            logger.info("Ecriture Item numero {} present dans les deux flux", element.getItemNumber());

            fillDppTypeWithDifference(sheet, styles, index, element, elementB, row);
            fillDppThirdWithDifference(sheet, styles, index, element, elementB, row);
            fillDppSubThirdWithDifference(sheet, styles, index, element, elementB, row);
            fillCustomerTypeWithDifference(sheet, styles, index, element, elementB, row);
            fillCustomerThirdWithDifference(sheet, styles, index, element, elementB, row);
            fillCustomerSubTirdWithDifference(sheet, styles, index, element, elementB, row);
            fillFinalTypeWithDifference(sheet, styles, index, element, elementB, row);
            fillFinalThirdWithDifference(sheet, styles, index, element, elementB, row);
            fillFinalSubThirdWithDifference(sheet, styles, index, element, elementB, row);
            fillItemNumberWithDifference(sheet, styles, index, element, elementB, row);
            fillCurrentWeekWithDifference(sheet, styles, index, element, elementB, row);
            fillCurrentYearWithDifference(sheet, styles, index, element, elementB, row);
            fillReplenishmentWithDifference(sheet, styles, index, element, elementB, row);
            fillImplementationWithDifference(sheet, styles, index, element, elementB, row);
        } else {
            logger.info("Ecriture Item numero {} present uniquement dans le flux d'origine", element.getItemNumber());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 0, 1, element.getDppType());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 2, 3, element.getDppThird());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 4, 5, element.getDppSubThird());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 6, 7, element.getCustomerType());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 8, 9, element.getCustomerThird());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 10, 11, element.getCustomerSubTihrd());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 12, 13, element.getFinalType());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 14, 15, element.getFinalThird());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 16, 17, element.getFinalSubThird());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 18, 19, element.getItemNumber());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 20, 21, element.getCurrentWeek());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 22, 23, element.getCurrentYear());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 24, 25, element.getReplenishment());
            fillMergedCell(sheet, styles.get(Color.RED), index, row, 26, 27, element.getImplementation());
        }
    }

    private static void fillDppTypeWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getDppType().equals(elementB.getDppType())) {
            fillCell(styles.get(Color.YELLOW), row, 0, element.getDppType());
            fillCell(styles.get(Color.YELLOW), row, 1, elementB.getDppType());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 0, 1, element.getDppType());
        }
    }

    private static void fillDppThirdWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getDppThird().equals(elementB.getDppThird())) {
            fillCell(styles.get(Color.YELLOW), row, 2, element.getDppThird());
            fillCell(styles.get(Color.YELLOW), row, 3, elementB.getDppThird());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 2, 3, element.getDppThird());
        }
    }

    private static void fillDppSubThirdWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getDppSubThird().equals(elementB.getDppSubThird())) {
            fillCell(styles.get(Color.YELLOW), row, 4, element.getDppSubThird());
            fillCell(styles.get(Color.YELLOW), row, 5, elementB.getDppSubThird());

        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 4, 5, element.getDppSubThird());
        }
    }

    private static void fillCustomerTypeWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getCustomerType().equals(elementB.getCustomerType())) {
            fillCell(styles.get(Color.YELLOW), row, 6, element.getCustomerType());
            fillCell(styles.get(Color.YELLOW), row, 7, elementB.getCustomerType());

        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 6, 7, element.getCustomerType());
        }
    }

    private static void fillCustomerThirdWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getCustomerThird().equals(elementB.getCustomerThird())) {
            fillCell(styles.get(Color.YELLOW), row, 8, element.getCustomerThird());
            fillCell(styles.get(Color.YELLOW), row, 9, elementB.getCustomerThird());

        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 8, 9, element.getCustomerThird());
        }
    }

    private static void fillCustomerSubTirdWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getCustomerSubTihrd().equals(elementB.getCustomerSubTihrd())) {
            fillCell(styles.get(Color.YELLOW), row, 10, element.getCustomerSubTihrd());
            fillCell(styles.get(Color.YELLOW), row, 11, elementB.getCustomerSubTihrd());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 10, 11, element.getCustomerSubTihrd());
        }
    }

    private static void fillFinalTypeWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getFinalType().equals(elementB.getFinalType())) {
            fillCell(styles.get(Color.YELLOW), row, 12, element.getFinalType());
            fillCell(styles.get(Color.YELLOW), row, 13, elementB.getFinalType());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 12, 13, element.getFinalType());
        }
    }

    private static void fillFinalThirdWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getFinalThird().equals(elementB.getFinalThird())) {
            fillCell(styles.get(Color.YELLOW), row, 14, element.getFinalThird());
            fillCell(styles.get(Color.YELLOW), row, 15, elementB.getFinalThird());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 14, 15, element.getFinalThird());
        }
    }

    private static void fillFinalSubThirdWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getFinalSubThird().equals(elementB.getFinalSubThird())) {
            fillCell(styles.get(Color.YELLOW), row, 16, element.getFinalSubThird());
            fillCell(styles.get(Color.YELLOW), row, 17, elementB.getFinalSubThird());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 16, 17, element.getFinalSubThird());
        }
    }

    private static void fillItemNumberWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getItemNumber().equals(elementB.getItemNumber())) {
            fillCell(styles.get(Color.YELLOW), row, 18, element.getItemNumber());
            fillCell(styles.get(Color.YELLOW), row, 19, elementB.getItemNumber());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 18, 19, element.getItemNumber());
        }
    }

    private static void fillCurrentWeekWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getCurrentWeek().equals(elementB.getCurrentWeek())) {
            fillCell(styles.get(Color.YELLOW), row, 20, element.getCurrentWeek());
            fillCell(styles.get(Color.YELLOW), row, 21, elementB.getCurrentWeek());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 20, 21, element.getCurrentWeek());
        }
    }

    private static void fillCurrentYearWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getCurrentYear().equals(elementB.getCurrentYear())) {
            fillCell(styles.get(Color.YELLOW), row, 22, element.getCurrentYear());
            fillCell(styles.get(Color.YELLOW), row, 23, elementB.getCurrentYear());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 22, 23, element.getCurrentYear());
        }
    }

    private static void fillReplenishmentWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getReplenishment().equals(elementB.getReplenishment())) {
            fillCell(styles.get(Color.YELLOW), row, 24, element.getReplenishment());
            fillCell(styles.get(Color.YELLOW), row, 25, elementB.getReplenishment());
        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 24, 25, element.getReplenishment());
        }
    }

    private static void fillImplementationWithDifference(Sheet sheet, EnumMap<Color, CellStyle> styles, int index, Element element, Element elementB, Row row) {
        if (!element.getImplementation().equals(elementB.getImplementation())) {
            fillCell(styles.get(Color.YELLOW), row, 26, element.getImplementation());
            fillCell(styles.get(Color.YELLOW), row, 27, elementB.getImplementation());

        } else {
            fillMergedCell(sheet, styles.get(Color.GREEN), index, row, 26, 27, element.getImplementation());
        }
    }

    private static void fillCell(CellStyle style, Row row, int i, String dppType) {
        Cell originalDppType = row.createCell(i);
        originalDppType.setCellValue(dppType);
        originalDppType.setCellStyle(style);
    }

    private static void fillMergedCell(Sheet sheet, CellStyle style, int index, Row row, int i, int i2, String implementation) {
        sheet.addMergedRegion(new CellRangeAddress(index, index, i, i2));
        fillCell(style, row, i, implementation);
        Cell emptyCell = row.createCell(i2);
        emptyCell.setCellStyle(style);
    }

    private static void saveExcelFile(Workbook workbook, String fileName) throws IOException {
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        
        String fileLocation = path.substring(0, path.length() - 1) + fileName;

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();
    }
}
