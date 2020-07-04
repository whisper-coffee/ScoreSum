package ssp;

import java.io.File;
import java.io.FileFilter;

public class FileFilterByXls implements FileFilter {
	
	private String suffix;
	public FileFilterByXls(String suffix) {
		super();
		this.suffix = suffix;
	}
	@Override
	public boolean accept(File pathname) {
		return pathname.getName().endsWith(suffix);
	}

}
