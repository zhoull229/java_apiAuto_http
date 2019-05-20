package com.wd.api.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.wd.api.cases.Case;
import com.wd.api.cases.Rest;

/**����Excel
 * @author Lenovo
 *
 */
public class ExcelUtil {
	
/**����Excel�ı������ݣ���װΪ����
 * @param <T>:����,֧�ִ��벻ͬ���͵Ķ���
 * @param clazz 
 * @param excelPath��Excel�ļ������·��
 * @param sheetName��Excel����
 */
public static <T> void lode(String excelPath, String sheetName, Class<T> clazz) {
	try {
		// ����workbook����
		Workbook workbook=WorkbookFactory.create(new File(excelPath));
		Sheet sheet=workbook.getSheet(sheetName);
		//��ȡ��һ��
		Row titleRow=sheet.getRow(0);
		//��ȡ�ж�����(�õ������к�),������ֵ��1
		int lastCellNum=titleRow.getLastCellNum();
		//����һ���ַ������飬������
		String [] fields=new String[lastCellNum];
		//ѭ������ÿһ��
		for (int i = 0; i < lastCellNum; i++) {
			//������������ȡ��Ӧ����
			Cell cell=titleRow.getCell(i,MissingCellPolicy.CREATE_NULL_AS_BLANK);
			//�����е�����Ϊ�ַ���
			cell.setCellType(CellType.STRING);
			//��ȡ�е�ֵ�����⣩
			String title=cell.getStringCellValue();
			//��ȡǰ���Ӣ�Ĳ���
			title=title.substring(0, title.indexOf("("));
			//�����ݷŵ�����
			fields[i]=title;
		
		}
		//��ȡ����,�õ�����������ֵ
		int lastRowIndex=sheet.getLastRowNum();
		//ѭ������ÿһ��������
		for (int i = 1; i <=lastRowIndex; i++) {
			//ͨ���ֽ��봴��case��Ķ���
			Object obj=clazz.newInstance();
			//�õ�һ��������
			Row dataRow=sheet.getRow(i);
			if (dataRow==null||isEmptyRow(dataRow)) {
				continue;
			}
			//�õ����������ϵ�ÿһ��,�����ݷ�װ��cs����
			for (int j = 0; j < lastCellNum; j++) {
				Cell cell=dataRow.getCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellType(CellType.STRING);
				String value=cell.getStringCellValue();
				//��ȡҪ����ķ�����
				String methodName="set"+fields[j];
				//��ȡҪ����ķ�������
				Method method=clazz.getMethod(methodName, String.class);
				//��ɷ������
				method.invoke(obj, value);

			}
			if (obj instanceof Case) {//instanceof�ж϶������͵��﷨
				Case cs=(Case) obj;
				//�Ѷ�����ӵ�����
				CaseUtil.cases.add(cs);
			}else if (obj instanceof Rest) {
				Rest rest=(Rest) obj;
				//�Ѷ�����ӵ�����
				RestUtil.rests.add(rest);
			}
			
		}
		
		
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} 
	
}

/**�п�
 * @param dataRow
 * @return
 */
private static boolean isEmptyRow(Row dataRow) {
	
	int lastCellNum =dataRow.getLastCellNum();
	for (int i = 0; i < lastCellNum; i++) {
		Cell cell=dataRow.getCell(i,MissingCellPolicy.CREATE_NULL_AS_BLANK);
		cell.setCellType(CellType.STRING);
		String value=cell.getStringCellValue();
		if (value!=null&&value.trim().length()>0) {
			return false;
		}
	}
	return true;
}


}
