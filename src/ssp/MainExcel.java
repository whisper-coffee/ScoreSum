package ssp;

public class MainExcel {
	/***
	 * ʹ��֮ǰ������������������ȫ������
	 * @param dir ���˱�����Ŀ¼
	 * @param file ������λ��
	 * @param xm_column ������
	 * @param zp_column ������
	 * @param jbsz_column ����������
	 * @param nlcj_column �����ɼ���
	 * @param fjf_column ���ӷ���
	 */
	public MainExcel(String dir, String file, int xm_column, int jbsz_column, int nlcj_column, int fjf_column, int zp_column) {
		new ExcelScore(dir, file, zp_column, jbsz_column, nlcj_column, fjf_column, xm_column);
	}
}
