package in.ainvnt.excel.parsing;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import in.ainvnt.excel.parsing.enums.ExcelSection;
import in.ainvnt.excel.parsing.model.ExcelField;
import in.ainvnt.excel.parsing.model.Order;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.fasterxml.jackson.databind.exc.InvalidFormatException;
import in.ainvnt.excel.parsing.mapper.ExcelFieldMapper;
import in.ainvnt.excel.parsing.model.Profit;
import in.ainvnt.excel.parsing.reader.ExcelFileReader;

public class ReadExcelGenericWay {

	public static void main(String[] args) throws InvalidFormatException {
		try {
			Workbook workbook = ExcelFileReader.readExcel("E:\\workspace\\projects\\Generic-Excel-Reading-POI\\order-profit.xlsx");
			Sheet sheet = workbook.getSheetAt(0);
			
			Map<String, List<ExcelField[]>> excelRowValuesMap = ExcelFileReader.getExcelRowValues(sheet);
			
			excelRowValuesMap.forEach((section, rows) -> {
				System.out.println(section);
				boolean headerPrint = true;
				for (ExcelField[] evc : rows) {
					if (headerPrint) {
						for (int j = 0; j < evc.length; j++) {
							System.out.print(evc[j].getExcelHeader() + "\t");
						}
						headerPrint = false;
					}
					for (int j = 0; j < evc.length; j++) {
						System.out.print(evc[j].getExcelValue() + "\t");
					}
					System.out.println();
				}
				System.out.println();
			});
			
			List<Order> orders = ExcelFieldMapper.getPojos(excelRowValuesMap.get(ExcelSection.ORDERS.getValue()),
					Order.class);

			List<Profit> profits = ExcelFieldMapper.getPojos(excelRowValuesMap.get(ExcelSection.PROFIT.getValue()),
					Profit.class);

			/*
			 * orders.forEach(o -> { System.out.println(o.getItem()); });
			 * 
			 * profits.forEach(p -> { System.out.println(p.getProfit());
			 * System.out.println(p.getDate()); });
			 */
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}
	}

}
