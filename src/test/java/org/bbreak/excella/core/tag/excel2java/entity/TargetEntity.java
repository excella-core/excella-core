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

package org.bbreak.excella.core.tag.excel2java.entity;

import java.math.BigDecimal;
import java.util.Date;

/**
 * テスト用エンティティ
 * 
 * @since 1.0
 */
public class TargetEntity {

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
