/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: WorkFactDetailData.java 2 2009-05-08 07:39:20Z yuta-takahashi $
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
package org.bbreak.excella.samples;

import java.util.Date;

/**
 * 勤務表の明細部分を保持するエンティティ
 * 
 * @since 1.0
 */
public class WorkFactDetailData {
    
    /** 日付 */
    private Date date;

    /** 出社時間 */
    private Double workStartHour;

    /** 退社時間 */
    private Double workEndHour;

    /** 作業内容 */
    private String workContent;

    public Date getDate() {
        return date;
    }

    public void setDate( Date date) {
        this.date = date;
    }

    public Double getWorkStartHour() {
        return workStartHour;
    }

    public void setWorkStartHour( Double workStartHour) {
        this.workStartHour = workStartHour;
    }

    public Double getWorkEndHour() {
        return workEndHour;
    }

    public void setWorkEndHour( Double workEndHour) {
        this.workEndHour = workEndHour;
    }

    public String getWorkContent() {
        return workContent;
    }

    public void setWorkContent( String workContent) {
        this.workContent = workContent;
    }

    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append( "[date=" + date);
        builder.append( ",workStartHour=" + workStartHour);
        builder.append( ",workEndHour=" + workEndHour);
        builder.append( ",workContent=" + workContent + "]");
        return builder.toString();
    }
}
