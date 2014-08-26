package org.bbreak.excella.core;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Cellの情報を保持するクラス
 * 
 * POI3.9より、xlsxで、Cellが削除されてRow内にCellがない状態になった時に、
 * あらかじめ取っておいたCellにアクセスしようとすると
 * XmlValueDisconnectedExceptionが発生する
 * 回避策として、このCellCloneオブジェクトに情報を保持しておくことができる
 * 比較時の関数などを以前の実装のまま使えるようにCellインターフェースを継承している
 * 情報を参照するためのオブジェクトなので、set系の関数を使用しようとすると
 * IllegalStateExceptionが発生する
 * 
 */
public class CellClone implements Cell {

	private int rowIndex;
	
	private int columnIndex;

	private CellStyle cellStyle;
	
	private int cellType;

	private Comment cellComment;
	
	private Row row;

	private Sheet sheet;
	
	private Hyperlink hyperlink;
	
	private boolean booleanCellValue;

	private String cellFormula;
	
	private Date dateCellValue;
	
	private byte errorCellValue;
	
	private double numericCellValue;
	
	private RichTextString richStringCellValue;

	/**
	 * コンストラクタ
	 * 
	 * @param cell
	 */
	public CellClone(Cell cell) {
		this.rowIndex = cell.getRowIndex();
		this.columnIndex = cell.getColumnIndex();
		this.cellStyle = cell.getCellStyle();
		this.cellType = cell.getCellType();
		this.cellComment = cell.getCellComment();
		this.row = cell.getRow();
		this.sheet = cell.getSheet();
		this.hyperlink = cell.getHyperlink();

		switch(this.cellType) {
		case CELL_TYPE_NUMERIC:
		case CELL_TYPE_FORMULA:
		case CELL_TYPE_BLANK:
			this.numericCellValue = cell.getNumericCellValue();
			this.dateCellValue = cell.getDateCellValue();
		}

		switch(this.cellType) {
		case CELL_TYPE_STRING:
		case CELL_TYPE_FORMULA:
		case CELL_TYPE_BLANK:
			this.richStringCellValue = cell.getRichStringCellValue();
		}

		switch(this.cellType) {
		case CELL_TYPE_FORMULA:
		case CELL_TYPE_BLANK:
		case CELL_TYPE_BOOLEAN:
			this.booleanCellValue = cell.getBooleanCellValue();
		}


		switch(this.cellType) {
		case CELL_TYPE_FORMULA:
			this.cellFormula = cell.getCellFormula();
		}
		
		switch(this.cellType) {
		case CELL_TYPE_ERROR:
			errorCellValue = cell.getErrorCellValue();
		}
	}

	public CellRangeAddress getArrayFormulaRange() {
		throw new IllegalStateException("CellClone is not support getArrayFormulaRange().");
	}

	public boolean getBooleanCellValue() {
		return booleanCellValue;
	}

	public int getCachedFormulaResultType() {
		throw new IllegalStateException("CellClone is not support getCachedFormulaResultType().");
	}

	public Comment getCellComment() {
		return cellComment;
	}

	public String getCellFormula() {
		return cellFormula;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public int getCellType() {
		return cellType;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public Date getDateCellValue() {
		return dateCellValue;
	}

	public byte getErrorCellValue() {
		return errorCellValue;
	}

	public Hyperlink getHyperlink() {
		return hyperlink;
	}

	public double getNumericCellValue() {
		return numericCellValue;
	}

	public RichTextString getRichStringCellValue() {
		return richStringCellValue;
	}

	public Row getRow() {
		return row;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public String getStringCellValue() {
        RichTextString str = getRichStringCellValue();
        return str == null ? null : str.getString();
	}

	public boolean isPartOfArrayFormulaGroup() {
		throw new IllegalStateException("CellClone is not support isPartOfArrayFormulaGroup().");
	}

	public void removeCellComment() {
		throw new IllegalStateException("CellClone is not support removeCellComment().");
	}

	public void setAsActiveCell() {
		throw new IllegalStateException("CellClone is not support setAsActiveCell().");
	}

	public void setCellComment(Comment comment) {
		throw new IllegalStateException("CellClone is not support setCellComment(Comment comment).");
	}

	public void setCellErrorValue(byte value) {
		throw new IllegalStateException("CellClone is not support setCellErrorValue(byte value).");
	}

	public void setCellFormula(String formula) throws FormulaParseException {
		throw new IllegalStateException("CellClone is not support setCellFormula(String formula).");
	}

	public void setCellStyle(CellStyle style) {
		throw new IllegalStateException("CellClone is not support setCellStyle(CellStyle style).");
	}

	public void setCellType(int cellType) {
		throw new IllegalStateException("CellClone is not support setCellType(int cellType).");
	}

	public void setCellValue(double value) {
		throw new IllegalStateException("CellClone is not support setCellValue(double value).");
	}

	public void setCellValue(Date value) {
		throw new IllegalStateException("CellClone is not support setCellValue(Date value).");
	}

	public void setCellValue(Calendar value) {
		throw new IllegalStateException("CellClone is not support setCellValue(Calendar value).");
	}

	public void setCellValue(RichTextString value) {
		throw new IllegalStateException("CellClone is not support setCellValue(RichTextString value).");
	}

	public void setCellValue(String value) {
		throw new IllegalStateException("CellClone is not support setCellValue(String value).");
	}

	public void setCellValue(boolean value) {
		throw new IllegalStateException("CellClone is not support setCellValue(boolean value).");
	}

	public void setHyperlink(Hyperlink link) {
		throw new IllegalStateException("CellClone is not support setHyperlink(Hyperlink link).");
	}

}
