package com.easy.excel;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * Java 简单的Excel 操作类 主要基于Jxl
 * 
 * @author JonkMing
 * 
 */
public class JEasyExcel {

	private WritableWorkbook workbook = null;

	private WritableSheet sheet = null;

	private boolean isFileExists = false;

	private File excelFile = null, copyFile = null;

	/**
	 * 打开excel 如果存在则打开，不存在则创建并打开
	 * 
	 * @param file
	 * @return
	 */
	public boolean open(File file) {
		if (null != sheet)
			colseExcel();
		try {
			// 由于jxl 打开之后文件是0 kb 所以如果该文件存在
			if (file.exists() && file.length() < 10) {
				// 如果文件大小小于 10字节 那么肯定是保存错误的 先删除 然后再创建
				file.delete();
				workbook = Workbook.createWorkbook(file);
				if (null == workbook)
					return false;
			} else if (file.exists()) {
				// 如果文件存在，为了防止被破坏，则先创建一个副本出来 然后编辑副本。最后如果没错的话将副本替换了。
				excelFile = file;
				FileUtils.copyFile(
						excelFile.getAbsolutePath(),
						excelFile.getAbsolutePath().substring(0,
								excelFile.getAbsolutePath().lastIndexOf("/"))
								+ "/copyExcel.xls");
				copyFile = new File(excelFile.getAbsolutePath().substring(0,
						excelFile.getAbsolutePath().lastIndexOf("/"))
						+ "/copyExcel.xls");
				isFileExists = true;
				Workbook rwb = Workbook.getWorkbook(copyFile);
				workbook = Workbook.createWorkbook(copyFile, rwb);
				if (null == workbook)
					return false;
			} else {

				workbook = Workbook.createWorkbook(file);
			}
			// workbook.write();
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return false;
		}

		return true;
	}

	/**
	 * 获取所有的Sheet 也就是工作表
	 * 
	 * @return 如果为空 则返回null
	 */
	public String[] getSheet() {
		// 先判断是否==null ，然后判断数组大小是否为0
		return workbook.getSheetNames() == null ? null : workbook
				.getSheetNames().length == 0 ? null : workbook.getSheetNames();
	}

	/**
	 * 创建sheet 这个方法会在所有的Sheet 后面添加一个Sheet
	 * 
	 * @param sheetName
	 *            名称
	 * @return
	 */
	public boolean createSheet(String sheetName) {

		try {
			sheet = workbook.createSheet(sheetName, getSheet() == null ? 0
					: getSheet().length);// 第一个参数为工作簿的名称，第二个参数为页数
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	 * 创建Sheet
	 * 
	 * @param sheetName
	 *            sheetName
	 * @param sheetIndex
	 *            sheetIndex坐标
	 * @return false 创建失败 true 创建成功
	 */
	public boolean createSheet(String sheetName, int sheetIndex) {

		try {
			sheet = workbook.createSheet(sheetName, sheetIndex);// 第一个参数为工作簿的名称，第二个参数为页数

		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	 * 写数据到Excel 报表
	 * 
	 * @param sheetName
	 *            sheet名称
	 * @param col
	 *            横坐标
	 * @param row
	 *            纵坐标
	 * @param value
	 *            值
	 * @return false 写入失败 true 写入成功
	 */
	public boolean writeDate(String sheetName, int col, int row, String value) {
		sheet = workbook.getSheet(sheetName);
		// sheet.
		try {
			sheet.setColumnView(col, 20); // 设置宽度

			sheet.addCell(new Label(col, row, value));
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	 * 写数据到报表
	 * 
	 * @param sheetIndex
	 *            sheet坐标
	 * @param col
	 *            横坐标
	 * @param row
	 *            纵坐标
	 * @param value
	 *            值
	 * @return false 写入成功 true 写入失败
	 */
	public boolean writeDate(int sheetIndex, int col, int row, String value) {
		sheet = workbook.getSheet(sheetIndex);
		try {
			sheet.setColumnView(col, 20); // 设置宽度
			sheet.addCell(new Label(col, row, value));
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	 * 写入数据到Excel 报表
	 * 
	 * @param sheetName
	 *            sheetName
	 * @param col
	 *            纵坐标
	 * @param row
	 *            横坐标
	 * @param value
	 *            值
	 * @param colour
	 *            背景颜色
	 * @return
	 */
	public boolean writeDate(String sheetName, int col, int row, String value,
			Colour colour) {
		sheet = workbook.getSheet(sheetName);
		// sheet.
		try {
			WritableCellFormat format = new WritableCellFormat();
			format.setBackground(colour);
			sheet.addCell(new Label(col, row, value, format));
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	 * 关闭Excel 报表
	 * 
	 * @return false 关闭失败 true 关闭成功
	 */
	public boolean colseExcel() {
		try {
			workbook.write();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			workbook = null;
			return false;
		} catch (WriteException e) {
			e.printStackTrace();
			workbook = null;
			return false;
		}
		// 如果关闭成功 没有保存则判断是否打开的副本
		if (isFileExists) {
			// 如果打开的是副本则覆盖过去
			if (excelFile.delete()) {
				// 如果删除成功
				FileUtils.copyFile(copyFile.getAbsolutePath(),
						excelFile.getAbsolutePath());
				copyFile.delete();
			}
		}
		workbook = null;
		return true;
	}

	/**
	 * 判断Sheet是否存在
	 * 
	 * @param sheetName
	 *            sheetName
	 * @return false 不存在 true 存在
	 */
	public boolean isSheetExist(String sheetName) {
		try {
			WritableSheet writableSheet = workbook.getSheet(sheetName);
			if (writableSheet == null)
				return false;
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return false;
		}
		return true;

	}

	/**
	 * 判断Sheet 的第几行0坐标是空的
	 * 
	 * @param sheetName
	 * @return
	 */
	public int isSheetColisNull(String sheetName) {
		sheet = workbook.getSheet(sheetName);
		int col = 0;
		Cell cell = null;
		while (true) {
			try {
				cell = sheet.getCell(col, 0);
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				return col;
			}

			if (cell.getContents() == null || cell.getContents() == "")
				break;
			col++;
		}
		return col;
	}

	/**
	 * 判断Sheet 的第几行0坐标是空的
	 * 
	 * @param sheetName
	 * @return
	 */
	public int isSheetColisNull(String sheetName, int row) {
		sheet = workbook.getSheet(sheetName);
		int col = 0;
		Cell cell = null;
		while (true) {
			try {
				cell = sheet.getCell(col, row);

			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				return col;
			}
			col++;
			if (cell == null || cell.getContents().length() <= 0)
				return --col;
			if (cell.getContents() != null || cell.getContents() != "")
				continue;
			else
				return --col;

		}

	}

	/**
	 * 获取指定坐标的值
	 * 
	 * @param sheetName
	 * @param Row
	 * @param Col
	 * @return
	 */
	public String getCallValue(String sheetName, int Row, int Col) {
		sheet = workbook.getSheet(sheetName);
		Cell cell = null;

		try {
			cell = sheet.getCell(Col, Row);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return null;
		}

		if (cell.getContents() == null || cell.getContents() == "")
			return null;

		return cell.getContents();
	}

	/**
	 * 判断Sheet 的第几行0坐标是空的
	 * 
	 * @param sheetName
	 * @return
	 */
	public int isSheetRowisNull(String sheetName) {
		sheet = workbook.getSheet(sheetName);
		int row = 0;
		Cell cell = null;
		while (true) {
			try {
				cell = sheet.getCell(0, row);
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				return row;
			}

			if (cell.getContents() == null || cell.getContents() == "")
				break;
			row++;
		}
		return row;
	}

	/**
	 * 设置工作簿密码
	 * 
	 * @param password
	 *            密码
	 * @return
	 */
	public void setWorkPassword(String password) {
		if (null != sheet)
			sheet.getSettings().setPassword(password);
	}

	/**
	 * 设置Sheet
	 * 
	 * @param index
	 * @return
	 */
	public void setSheet(int index) {
		sheet = workbook.getSheet(index);
	}

	/**
	 * 设置Sheet
	 * 
	 * @param sheetName
	 * @return
	 */
	public void setSheet(String sheetName) {
		sheet = workbook.getSheet(sheetName);
	}

	/**
	 * 设置Sheet
	 * 
	 * @param index
	 * @param passwd
	 * @return
	 */
	public void setSheet(int index, String passwd) {
		sheet = workbook.getSheet(index);
		sheet.getSettings().setPassword(passwd);
	}

	/**
	 * 设置Sheet
	 * 
	 * @param index
	 * @param passwd
	 * @return
	 */
	public void setSheet(String sheetName, String passwd) {
		sheet = workbook.getSheet(sheetName);
		sheet.getSettings().setPassword(passwd);
	}

}
