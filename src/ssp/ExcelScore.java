package ssp;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelScore {
	
	private int xm_column;
	private int zp_column;
	private int jbsz_column;
	private int nlcj_column;
	private int fjf_column;
	
	public ExcelScore(String dir, String file, int zp_column, int jbsz_column, int nlcj_column, int fjf_column, int xm_column) {
		try {
			this.xm_column = xm_column;
			this.zp_column = zp_column;
			this.jbsz_column = jbsz_column;
			this.nlcj_column = nlcj_column;
			this.fjf_column = fjf_column;
			startRun(dir,file);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * startRun:进行执行
	 * */
	private void startRun(String strdir, String strfile) throws IOException, InvalidFormatException {
		File dir = new File(strdir);
		File file = new File(strfile);
		Map<String,String> map = new TreeMap<String, String>();
		List<File> list = fillList(dir,".xls");

		fileInput(dir,list,map,65,9);//(66,10)：个人表里综评成绩位置（行列减一），以下三项同理
		System.out.println("综评成绩提取完毕");
		writeToXlsx(map,zp_column,file);
		System.out.println("综评写入完毕");
		
		fileInput(dir,list,map,6,9);//个人表里基本素质成绩位置
		System.out.println("基本素质成绩提取完毕");
		writeToXlsx(map,jbsz_column,file);
		System.out.println("基本素质成绩写入完毕");

		fileInput(dir,list,map,35,9);//个人表里能力成绩位置
		System.out.println("能力成绩提取完毕");
		writeToXlsx(map,nlcj_column,file);
		System.out.println("能力成绩写入完毕");
		
		fileInput(dir,list,map,59,9);//个人表里附加分位置
		System.out.println("附加分提取完毕");
		writeToXlsx(map,fjf_column,file);
		System.out.println("附加分写入完毕");
		System.out.println("数据写入完成");
		//JOptionPane.showMessageDialog(null, "数据写入完成");
	}
	/**
	 * fillList 调用getXlsFile方法，获取.xls文件
	 * @param dir 文件目录
	 * @param suffix 个人文件列表
	 * */
	private List<File> fillList(File dir, String suffix) {
		List<File> list = new ArrayList<File>();
		FileFilter filter = new FileFilterByXls(suffix);
		getXlsFile(dir,list,filter);//获取所有的xls文件存入 list
		return list;
	}
	/**
	 * getXlsFile 递归获取xls文件并存入list
	 * @param dir 文件目录
	 * @param list 个人文件列表
	 * @param filter 文件过滤器
	 * */
	private void getXlsFile(File dir, List<File> list, FileFilter filter) {
		File[] files = dir.listFiles();
		for (File file : files) {
			if(file.isDirectory())
				getXlsFile(file, list, filter);
			else {
				if(filter.accept(file)) {
					list.add(file);
				}
			}
		}
	}
	/**
	 * fileInput 提取成绩
	 * @param dir 文件目录
	 * @param list 个人文件列表
	 * @param map 名字做键，提取出来的成绩做值
	 * @param row 个人文件目标行
	 * @param cell 个人文件目标单元格
	 * @exception IOException
	 * */
	private void fileInput(File dir, List<File> list, Map<String, String> map, int row, int cell) throws IOException {
		for (Iterator<File> it = list.iterator(); it.hasNext();) {
			File file = (File) it.next();
			FileInputStream fis = new FileInputStream(file);
			getExcelCell(fis,file,map,row,cell);
			fis.close();
		}
		System.out.println("共"+list.size()+"人");
	}
	/**
	 * getExcelCell 从个人表中提取成绩
	 * @param fis 文件输入流
	 * @param file 个人文件
	 * @param map 名字做键，提取出来的成绩做值
	 * @param row 个人文件目标行
	 * @param cell 个人文件目标单元格
	 * @exception IOException
	 * */
	private void getExcelCell(FileInputStream fis, File file, Map<String, String> map, int row, int cell) throws IOException {
		HSSFWorkbook hwb = new HSSFWorkbook(fis);//获取excel文件
		HSSFSheet sheet = hwb.getSheetAt(0);//获取sheet
		HSSFRow hrow = sheet.getRow(row);//获取行数
		HSSFCell hcell = hrow.getCell(cell);//获取单元格
		hcell.setCellType(CellType.STRING);//改变单元格类型以便取出
		String str = hcell.getStringCellValue();//得到成绩
		if(str=="")
			str = "0";
		double d = Double.parseDouble(str);
		if(d>20)
			str = String.format("%.4f", d);//综评格式为四位小数
		String name = "";
		name = new String(file.getName().getBytes()).substring(10, file.getName().length()-4);//从文件名获取人名存入map作为键
		map.put(name, str);
		hwb.close();
	}
	/**
	 * writeToXlsx 将成绩存入总表中对应学生的格内
	 * @param map 学生姓名及成绩
	 * @param cell 要填入的单元格
	 * @param file 总表文件
	 * @throws IOException
	 * @throws InvalidFormatException
	 * */
	private void writeToXlsx(Map<String, String> map, int cell, File file) throws IOException, InvalidFormatException {
		FileInputStream fis = new FileInputStream(file);
		Workbook xwb = new XSSFWorkbook(fis);
		fis.close();
		XSSFSheet sheet = (XSSFSheet) xwb.getSheetAt(0);
		Set<String> set = map.keySet();
		boolean flag;
		for (Iterator<String> setit = set.iterator(); setit.hasNext();) {
			flag = false;
			String name = (String)setit.next();
			Iterator<Row> rowit = sheet.rowIterator();
//			int i = 0;//计数器
			while(rowit.hasNext()) {
				XSSFRow row = (XSSFRow) rowit.next();
				XSSFCell namecell = (XSSFCell) row.getCell(xm_column);//姓名所在列位置
				XSSFCell scorecell = row.getCell(cell);
				try {
					namecell.setCellType(CellType.STRING);
				} catch (Exception e) {
					JOptionPane.showMessageDialog(null, "在写入数据时发生异常\n检查表列名、表格式、个人文件名等是否正确");
					System.exit(0);
				}
				String valuecell = namecell.getStringCellValue();
				if(valuecell.equals(name)) {
					scorecell.setCellValue(map.get(name));//根据姓名匹配写入成绩
					flag = true;
				}
				if (flag) {
					break;
				}
//				i++;
			}
		}
		FileOutputStream fos = new FileOutputStream(file);
		xwb.write(fos);
		fos.close();
		xwb.close();
	}
	/**
	 * printMap:打印map(测试用)
	 * @param map
	 * */
	@SuppressWarnings("unused")
	private void printMap(Map<String, String> map) {
		Set<String> set = map.keySet();
		for (Iterator<String> it = set.iterator(); it.hasNext();) {
			String name = (String)it.next();
			String socre = map.get(name);
			System.out.println(name+"---"+socre);
		}
	}
}
