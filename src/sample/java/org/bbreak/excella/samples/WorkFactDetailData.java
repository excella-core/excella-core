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
