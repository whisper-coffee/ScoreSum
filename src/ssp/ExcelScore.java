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
	 * startRun:����ִ��
	 * */
	private void startRun(String strdir, String strfile) throws IOException, InvalidFormatException {
		File dir = new File(strdir);
		File file = new File(strfile);
		Map<String,String> map = new TreeMap<String, String>();
		List<File> list = fillList(dir,".xls");

		fileInput(dir,list,map,65,9);//(66,10)�����˱��������ɼ�λ�ã����м�һ������������ͬ��
		System.out.println("�����ɼ���ȡ���");
		writeToXlsx(map,zp_column,file);
		System.out.println("����д�����");
		
		fileInput(dir,list,map,6,9);//���˱���������ʳɼ�λ��
		System.out.println("�������ʳɼ���ȡ���");
		writeToXlsx(map,jbsz_column,file);
		System.out.println("�������ʳɼ�д�����");

		fileInput(dir,list,map,35,9);//���˱��������ɼ�λ��
		System.out.println("�����ɼ���ȡ���");
		writeToXlsx(map,nlcj_column,file);
		System.out.println("�����ɼ�д�����");
		
		fileInput(dir,list,map,59,9);//���˱��︽�ӷ�λ��
		System.out.println("���ӷ���ȡ���");
		writeToXlsx(map,fjf_column,file);
		System.out.println("���ӷ�д�����");
		System.out.println("����д�����");
		//JOptionPane.showMessageDialog(null, "����д�����");
	}
	/**
	 * fillList ����getXlsFile��������ȡ.xls�ļ�
	 * @param dir �ļ�Ŀ¼
	 * @param suffix �����ļ��б�
	 * */
	private List<File> fillList(File dir, String suffix) {
		List<File> list = new ArrayList<File>();
		FileFilter filter = new FileFilterByXls(suffix);
		getXlsFile(dir,list,filter);//��ȡ���е�xls�ļ����� list
		return list;
	}
	/**
	 * getXlsFile �ݹ��ȡxls�ļ�������list
	 * @param dir �ļ�Ŀ¼
	 * @param list �����ļ��б�
	 * @param filter �ļ�������
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
	 * fileInput ��ȡ�ɼ�
	 * @param dir �ļ�Ŀ¼
	 * @param list �����ļ��б�
	 * @param map ������������ȡ�����ĳɼ���ֵ
	 * @param row �����ļ�Ŀ����
	 * @param cell �����ļ�Ŀ�굥Ԫ��
	 * @exception IOException
	 * */
	private void fileInput(File dir, List<File> list, Map<String, String> map, int row, int cell) throws IOException {
		for (Iterator<File> it = list.iterator(); it.hasNext();) {
			File file = (File) it.next();
			FileInputStream fis = new FileInputStream(file);
			getExcelCell(fis,file,map,row,cell);
			fis.close();
		}
		System.out.println("��"+list.size()+"��");
	}
	/**
	 * getExcelCell �Ӹ��˱�����ȡ�ɼ�
	 * @param fis �ļ�������
	 * @param file �����ļ�
	 * @param map ������������ȡ�����ĳɼ���ֵ
	 * @param row �����ļ�Ŀ����
	 * @param cell �����ļ�Ŀ�굥Ԫ��
	 * @exception IOException
	 * */
	private void getExcelCell(FileInputStream fis, File file, Map<String, String> map, int row, int cell) throws IOException {
		HSSFWorkbook hwb = new HSSFWorkbook(fis);//��ȡexcel�ļ�
		HSSFSheet sheet = hwb.getSheetAt(0);//��ȡsheet
		HSSFRow hrow = sheet.getRow(row);//��ȡ����
		HSSFCell hcell = hrow.getCell(cell);//��ȡ��Ԫ��
		hcell.setCellType(CellType.STRING);//�ı䵥Ԫ�������Ա�ȡ��
		String str = hcell.getStringCellValue();//�õ��ɼ�
		if(str=="")
			str = "0";
		double d = Double.parseDouble(str);
		if(d>20)
			str = String.format("%.4f", d);//������ʽΪ��λС��
		String name = "";
		name = new String(file.getName().getBytes()).substring(10, file.getName().length()-4);//���ļ�����ȡ��������map��Ϊ��
		map.put(name, str);
		hwb.close();
	}
	/**
	 * writeToXlsx ���ɼ������ܱ��ж�Ӧѧ���ĸ���
	 * @param map ѧ���������ɼ�
	 * @param cell Ҫ����ĵ�Ԫ��
	 * @param file �ܱ��ļ�
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
//			int i = 0;//������
			while(rowit.hasNext()) {
				XSSFRow row = (XSSFRow) rowit.next();
				XSSFCell namecell = (XSSFCell) row.getCell(xm_column);//����������λ��
				XSSFCell scorecell = row.getCell(cell);
				try {
					namecell.setCellType(CellType.STRING);
				} catch (Exception e) {
					JOptionPane.showMessageDialog(null, "��д������ʱ�����쳣\n�������������ʽ�������ļ������Ƿ���ȷ");
					System.exit(0);
				}
				String valuecell = namecell.getStringCellValue();
				if(valuecell.equals(name)) {
					scorecell.setCellValue(map.get(name));//��������ƥ��д��ɼ�
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
	 * printMap:��ӡmap(������)
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
