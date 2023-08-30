package me.rustynail.compareCells.controller;

import jakarta.annotation.Resource;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Executor;
import java.util.concurrent.TimeUnit;

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

		Workbook wb = null;
		Workbook fileBook1 = null;
		Workbook fileBook2 = null;


		if (file1 == null || file2 == null || file1.isEmpty() || file2.isEmpty()) {
			response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
			response.getWriter().write("没选到文件？");
			return;
		}

		try {
			fileBook1 = new XSSFWorkbook(file1.getInputStream());
		} catch (Exception e) {
			log.error(e.getMessage());
			fileBook1 = new HSSFWorkbook(file1.getInputStream());
		}

		try {
			fileBook2 = new XSSFWorkbook(file2.getInputStream());
		} catch (Exception e) {
			log.error(e.getMessage());
			fileBook2 = new HSSFWorkbook(file2.getInputStream());
		}


		wb = compareJob(fileBook1, fileBook2);


		response.setStatus(HttpStatus.OK.value());
		response.setHeader(HttpHeaders.CONTENT_DISPOSITION,
				"attachment;filename=" + URLEncoder.encode(file1.getOriginalFilename(), StandardCharsets.UTF_8));
		ByteArrayOutputStream bao = new ByteArrayOutputStream();
		wb.write(bao);
		bao.flush();
		byte[] buf = bao.toByteArray();
		response.getOutputStream().write(buf);
		response.getOutputStream().flush();
	}

	private Workbook compareJob(Workbook wb1, Workbook wb2) throws InterruptedException {

		int sheetNums1 = wb1.getNumberOfSheets();
		// 选大的

		CountDownLatch latch = new CountDownLatch(sheetNums1);
		for (var s = 0; s < sheetNums1; s++) {
			try {
				int finalS = s;
				compareJob(wb1, wb2, finalS);
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

	private void compareJob(Workbook wb1, Workbook wb2, int s) {
		// if sheet count < s
		if (wb1.getNumberOfSheets() <= s || wb2.getNumberOfSheets() <= s) {
			return;
		}


		var sheet1 = wb1.getSheetAt(s);
		var sheet2 = wb2.getSheetAt(s);

		if (sheet1 == null || sheet2 == null) {
			return;
		}

		// 假如名字不一样
		if (!sheet1.getSheetName().equals(sheet2.getSheetName())) {
			log.warn("Sheet 名字不一样: [{}, {}]", sheet1.getSheetName(), sheet2.getSheetName());
			if (sheet1 instanceof XSSFSheet) {
				((XSSFSheet) sheet1).setTabColor(new XSSFColor(new java.awt.Color(216, 191, 216), new DefaultIndexedColorMap()));
			}
		}
		compareSheets(s, sheet1, sheet2);
	}

	private void compareSheets(int s, Sheet sheet1, Sheet sheet2) {
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
			if (row1 == null) {
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

				if (cell1 == null && cell2 != null && !"".equals(cell2.toString().trim())) {
					Cell cell = row1.createCell(c);
					setCellColor(cell);
					continue;
				}

				if (cell1.toString().trim().equals(cell2.toString().trim())) {
					continue;
				} else {
					log.info(sheet1.getSheetName() + "--" + sheet1.getSheetName() + " - " + r + " -- " + c
							+ "[" + cell1 + " : "
							+ cell2 + "]");
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
		if (cellStyle instanceof HSSFCellStyle) {
			HSSFCellStyle style = (HSSFCellStyle) cell.getSheet().getWorkbook().createCellStyle();

			style.cloneStyleFrom(cellStyle);
			HSSFPalette customPalette = ((HSSFWorkbook) cell.getRow().getSheet().getWorkbook()).getCustomPalette();
			customPalette.setColorAtIndex(HSSFColor.HSSFColorPredefined.LEMON_CHIFFON.getIndex(), (byte) 111, (byte) 11, (byte) 111);
			HSSFColor similarColor = customPalette.findColor((byte) 111, (byte) 11, (byte) 111);
			style.setFillBackgroundColor(similarColor);
//			style.setFillBackgroundColor(similarColor);
			style.setFillPattern(FillPatternType.DIAMONDS);
			cell.setCellStyle(style);
		} else {
			cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(216, 191, 216), new DefaultIndexedColorMap()));
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell.setCellStyle(cellStyle);
		}
		// cellStyle.setFillForegroundColor(IndexedColors.VIOLET.index);
		;
		// cellStyle.setFillForegroundColor(IndexedColors.TEAL.getIndex());

	}
}
