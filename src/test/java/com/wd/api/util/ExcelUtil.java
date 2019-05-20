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

/**解析Excel
 * @author Lenovo
 *
 */
public class ExcelUtil {
	
/**解析Excel的表单的数据，封装为对象
 * @param <T>:泛型,支持传入不同类型的对象
 * @param clazz 
 * @param excelPath：Excel文件的相对路径
 * @param sheetName：Excel表单名
 */
public static <T> void lode(String excelPath, String sheetName, Class<T> clazz) {
	try {
		// 创建workbook对象
		Workbook workbook=WorkbookFactory.create(new File(excelPath));
		Sheet sheet=workbook.getSheet(sheetName);
		//获取第一行
		Row titleRow=sheet.getRow(0);
		//获取有多少列(得到的是列号),比索引值大1
		int lastCellNum=titleRow.getLastCellNum();
		//创建一个字符串数组，放数据
		String [] fields=new String[lastCellNum];
		//循环处理每一列
		for (int i = 0; i < lastCellNum; i++) {
			//根据列索引获取对应的列
			Cell cell=titleRow.getCell(i,MissingCellPolicy.CREATE_NULL_AS_BLANK);
			//设置列的类型为字符串
			cell.setCellType(CellType.STRING);
			//获取列的值（标题）
			String title=cell.getStringCellValue();
			//截取前面的英文部分
			title=title.substring(0, title.indexOf("("));
			//将数据放到数组
			fields[i]=title;
		
		}
		//获取行数,得到的是行索引值
		int lastRowIndex=sheet.getLastRowNum();
		//循环处理每一个数据行
		for (int i = 1; i <=lastRowIndex; i++) {
			//通过字节码创建case类的对象
			Object obj=clazz.newInstance();
			//拿到一个数据行
			Row dataRow=sheet.getRow(i);
			if (dataRow==null||isEmptyRow(dataRow)) {
				continue;
			}
			//拿到此数据行上的每一列,将数据封装到cs对象
			for (int j = 0; j < lastCellNum; j++) {
				Cell cell=dataRow.getCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellType(CellType.STRING);
				String value=cell.getStringCellValue();
				//获取要反射的方法名
				String methodName="set"+fields[j];
				//获取要反射的方法对象
				Method method=clazz.getMethod(methodName, String.class);
				//完成反射调用
				method.invoke(obj, value);

			}
			if (obj instanceof Case) {//instanceof判断对象类型的语法
				Case cs=(Case) obj;
				//把对象添加到集合
				CaseUtil.cases.add(cs);
			}else if (obj instanceof Rest) {
				Rest rest=(Rest) obj;
				//把对象添加到集合
				RestUtil.rests.add(rest);
			}
			
		}
		
		
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} 
	
}

/**判空
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
