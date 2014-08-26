/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: IllegalTargetEntity.java 2 2009-05-08 07:39:20Z yuta-takahashi $
 * $Revision: 2 $
 *
 * This file is part of ExCella Core.
 *
 * ExCella Core is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License version 3
 * only, as published by the Free Software Foundation.
 *
 * ExCella Core is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License version 3 for more details
 * (a copy is included in the COPYING.LESSER file that accompanied this code).
 *
 * You should have received a copy of the GNU Lesser General Public License
 * version 3 along with ExCella Core.  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
package org.bbreak.excella.core.tag.excel2java.entity;

import java.math.BigDecimal;
import java.util.Date;

/**
 * テスト用エンティティ
 * 
 * @since 1.0
 */
public class IllegalTargetEntity {

    /**
     * Integer型
     */
    private Integer numberInteger;

    /**
     * int型
     */
    private int numberInt;

    /**
     * Long型
     */
    private Long numberLong;

    /**
     * long型
     */
    private long numberlong;

    /**
     * Folat型
     */
    private Float numberFloat;

    /**
     * float型
     */
    private float numberfloat;

    /**
     * Double型
     */
    private Double numberDouble;

    /**
     * double型
     */
    private double numberdouble;

    /**
     * BigDecimal型
     */
    private BigDecimal numberDecimal;

    /**
     * Byte型
     */
    private Byte numberByte;

    /**
     * byte型
     */
    private byte numberbyte;

    /**
     * Boolean型
     */
    private Boolean valueBoolean;

    /**
     * boolean型
     */
    private boolean valueboolean;

    /**
     * Date型
     */
    private Date date;

    /**
     * String型
     */
    private String string;

    /**
     * ChildEntity型
     */
    private ChildEntity childEntity;

    /**
     * privateコンストラクタ
     */
    private IllegalTargetEntity() {
    }
    
    public Integer getNumberInteger() {
        return numberInteger;
    }

    public void setNumberInteger( Integer numberInteger) {
        this.numberInteger = numberInteger;
    }

    public int getNumberInt() {
        return numberInt;
    }

    public void setNumberInt( int numberInt) {
        this.numberInt = numberInt;
    }

    public Long getNumberLong() {
        return numberLong;
    }

    public void setNumberLong( Long numberLong) {
        this.numberLong = numberLong;
    }

    public long getNumberlong() {
        return numberlong;
    }

    public void setNumberlong( long numberlong) {
        this.numberlong = numberlong;
    }

    public Float getNumberFloat() {
        return numberFloat;
    }

    public void setNumberFloat( Float numberFloat) {
        this.numberFloat = numberFloat;
    }

    public float getNumberfloat() {
        return numberfloat;
    }

    public void setNumberfloat( float numberfloat) {
        this.numberfloat = numberfloat;
    }

    public Double getNumberDouble() {
        return numberDouble;
    }

    public void setNumberDouble( Double numberDouble) {
        this.numberDouble = numberDouble;
    }

    public double getNumberdouble() {
        return numberdouble;
    }

    public void setNumberdouble( double numberdouble) {
        this.numberdouble = numberdouble;
    }

    public BigDecimal getNumberDecimal() {
        return numberDecimal;
    }

    public void setNumberDecimal( BigDecimal numberDecimal) {
        this.numberDecimal = numberDecimal;
    }

    public Byte getNumberByte() {
        return numberByte;
    }

    public void setNumberByte( Byte numberByte) {
        this.numberByte = numberByte;
    }

    public byte getNumberbyte() {
        return numberbyte;
    }

    public void setNumberbyte( byte numberbyte) {
        this.numberbyte = numberbyte;
    }

    public Boolean getValueBoolean() {
        return valueBoolean;
    }

    public void setValueBoolean( Boolean valueBoolean) {
        this.valueBoolean = valueBoolean;
    }

    public boolean isValueboolean() {
        return valueboolean;
    }

    public void setValueboolean( boolean valueboolean) {
        this.valueboolean = valueboolean;
    }

    public Date getDate() {
        return date;
    }

    public void setDate( Date date) {
        this.date = date;
    }

    public String getString() {
        return string;
    }

    public void setString( String string) {
        this.string = string;
    }

    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append( "[numberInteger=" + numberInteger);
        builder.append( "][numberInt=" + numberInt);
        builder.append( "][numberLong=" + numberLong);
        builder.append( "][numberlong=" + numberlong);
        builder.append( "][numberFloat=" + numberFloat);
        builder.append( "][numberfloat=" + numberfloat);
        builder.append( "][numberDouble=" + numberDouble);
        builder.append( "][numberdouble=" + numberdouble);
        builder.append( "][numberDecimal=" + numberDecimal);
        builder.append( "][numberByte=" + numberByte);
        builder.append( "][numberbyte=" + numberbyte);
        builder.append( "][valueBoolean=" + valueBoolean);
        builder.append( "][valueboolean=" + valueboolean);
        builder.append( "][date=" + date);
        builder.append( "][string=" + string);
        builder.append( "]");

        return builder.toString();
    }

    public ChildEntity getChildEntity() {
        return childEntity;
    }

    public void setChildEntity( ChildEntity childEntity) {
        this.childEntity = childEntity;
    }
}
