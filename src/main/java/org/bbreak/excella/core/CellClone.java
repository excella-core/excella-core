/*-
 * #%L
 * excella-core
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.core;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
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
	
	private CellType cellTypeEnum;

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
		this.cellTypeEnum = cell.getCellTypeEnum();
		this.cellComment = cell.getCellComment();
		this.row = cell.getRow();
		this.sheet = cell.getSheet();
		this.hyperlink = cell.getHyperlink();

        switch ( this.cellTypeEnum) {
            case STRING:
                this.richStringCellValue = cell.getRichStringCellValue();
                break;
            case NUMERIC:
                this.numericCellValue = cell.getNumericCellValue();
                this.dateCellValue = cell.getDateCellValue();
                break;
            case FORMULA:
                this.numericCellValue = cell.getNumericCellValue();
                this.dateCellValue = cell.getDateCellValue();
                this.richStringCellValue = cell.getRichStringCellValue();
                this.booleanCellValue = cell.getBooleanCellValue();
                this.cellFormula = cell.getCellFormula();
                break;
            case BOOLEAN:
                this.booleanCellValue = cell.getBooleanCellValue();
                break;
            case BLANK:
                this.numericCellValue = cell.getNumericCellValue();
                this.dateCellValue = cell.getDateCellValue();
                this.richStringCellValue = cell.getRichStringCellValue();
                this.booleanCellValue = cell.getBooleanCellValue();
                break;
            case ERROR:
                errorCellValue = cell.getErrorCellValue();
                break;
            default:
                break;
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

	@Deprecated
	public int getCellType() {
		return cellTypeEnum.getCode();
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
	
	@Override
	public void removeHyperlink() {
        throw new IllegalStateException("CellClone is not support removeHyperlink().");
	}

    @Override
    public void setCellType( CellType cellType) {
        throw new IllegalStateException("CellClone is not support setCellType().");
    }

    @Override
    public CellType getCellTypeEnum() {
        return cellTypeEnum;
    }

    @Override
    public CellType getCachedFormulaResultTypeEnum() {
        throw new IllegalStateException("CellClone is not support getCachedFormulaResultTypeEnum().");
    }

    @Override
    public CellAddress getAddress() {
        return new CellAddress(this);
    }

}
