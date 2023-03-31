package me.rustynail.compareCells.controller;

import java.net.URLEncoder;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Executor;
import java.util.concurrent.TimeUnit;

import jakarta.annotation.Resource;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFColorScaleFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;

@Slf4j
@RestController
@RequestMapping("/api")
public class Api {

    @Resource(name = "queryTask")
    private Executor executor;

    /**
     * compare
     *
     * @param file1
     * @param file2
     */
    @PostMapping("/cmp")
    public void compare(
            @RequestParam("file1") MultipartFile file1,
            @RequestParam("file2") MultipartFile file2,
            HttpServletResponse response) throws Exception {
        Workbook wb = compareJob(new XSSFWorkbook(file1.getInputStream()), new XSSFWorkbook(file2.getInputStream()));

        response.setStatus(HttpStatus.OK.value());
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION,
                "attachment;filename*=UTF-8''" + URLEncoder.encode(file1.getOriginalFilename(), "utf-8"));
        wb.write(response.getOutputStream());
        response.getOutputStream().flush();
    }

    private Workbook compareJob(Workbook wb1, Workbook wb2) throws InterruptedException {

        int sheetNums1 = wb1.getNumberOfSheets();
        // 选大的

        CountDownLatch latch = new CountDownLatch(sheetNums1);
        for (var s = 0; s < sheetNums1; s++) {
            try {
                int finalS = s;
                compareSheets(wb1, wb2, finalS);
//                executor.execute(() -> compareSheets(wb1, wb2, finalS));
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                latch.countDown();
            }
        }
        latch.await(10, TimeUnit.MINUTES);

        return wb1;
    }

    private void compareSheets(Workbook wb1, Workbook wb2, int s) {
        var sheet1 = wb1.getSheetAt(s);
        var sheet2 = wb2.getSheetAt(s);

        if (sheet1 == null || sheet2 == null) {
            return;
        }
        log.info("处理->" + sheet1.getSheetName() + "[" + s + "]" + " -- " + sheet2.getSheetName());
        var rowNums1 = sheet1.getLastRowNum();

        for (int r = 0; r <= rowNums1; r++) {

            var row1 = sheet1.getRow(r);

            var row2 = sheet2.getRow(r);

            if (row1 == null && row2 == null) {
                log.info(sheet1.getSheetName() + "--" + r + " 全空行跳过");
                continue;
            }

            // 在新文件里有，旧文件没有
            if (row1 != null && row2 == null) {
                for (int c = 0; c < row1.getLastCellNum(); c++) {
                    var cell1 = row1.getCell(c);
                    if (cell1 == null) {
                        cell1 = row1.createCell(c);
                    }
                    setCellColor(cell1);
                }
                continue;
            }

            // 在新文件里没有，旧文件有
            if (row1 == null && row2 != null) {
                log.info(sheet1.getSheetName() + "--" + r + " 新文件空行跳过");
                continue;
            }

            var colNums1 = row1.getLastCellNum();

            for (int c = 0; c <= colNums1; c++) {

                var cell1 = row1.getCell(c);

                var cell2 = row2.getCell(c);

                if (null == cell1 && null == cell2) {
                    log.info(sheet1.getSheetName() + "--" + r + " -- " + c + " 空cell跳过");
                    continue;
                }

                if (null != cell1 && null == cell2) {
                    setCellColor(cell1);
                    continue;
                }

                if (cell1.toString().trim().equals(cell2.toString().trim())) {
                    continue;
                } else {
                    log.info(sheet1.getSheetName() + "--" + sheet1.getSheetName() + " - " + r + " -- " + c
                            + "[" + cell1.toString() + " : "
                            + cell2.toString() + "]");
                    setCellColor(cell1);
                }

            }
        }
    }

    private void setCellColor(Cell cell) {
        log.info("染色：" + cell.getSheet().getSheetName() + "- [" + cell.getRowIndex() + ","
                + cell.getColumnIndex() + "]");
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        } else {
            var cellStylea = cell.getSheet().getWorkbook().createCellStyle();
            cellStylea.cloneStyleFrom(cellStyle);
            cellStyle = cellStylea;
        }
        // #fff
        // cellStyle.setFillForegroundColor(IndexedColors.VIOLET.index);
        cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(216, 191, 216), new DefaultIndexedColorMap()));
        // cellStyle.setFillForegroundColor(IndexedColors.TEAL.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
    }
}
