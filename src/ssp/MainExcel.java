package ssp;

public class MainExcel {
	/***
	 * 使用之前必须在综评表内填入全班姓名
	 * @param dir 个人表所在目录
	 * @param file 综评表位置
	 * @param xm_column 姓名列
	 * @param zp_column 综评列
	 * @param jbsz_column 基本素质列
	 * @param nlcj_column 能力成绩列
	 * @param fjf_column 附加分列
	 */
	public MainExcel(String dir, String file, int xm_column, int jbsz_column, int nlcj_column, int fjf_column, int zp_column) {
		new ExcelScore(dir, file, zp_column, jbsz_column, nlcj_column, fjf_column, xm_column);
	}
}
