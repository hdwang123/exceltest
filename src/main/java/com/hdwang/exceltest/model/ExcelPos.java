package com.hdwang.exceltest.model;

/**
 * 表格位置
 *
 * @author wanghuidong
 * 时间： 2022/8/8 18:02
 */
public class ExcelPos {

    /**
     * 起始位置，例如：A1
     */
    private String startPos;

    /**
     * 结束位置，例如：A2
     */
    private String endPos;

    public ExcelPos(String startPos, String endPos) {
        this.startPos = startPos;
        this.endPos = endPos;
    }

    public String getStartPos() {
        return startPos;
    }

    public void setStartPos(String startPos) {
        this.startPos = startPos;
    }

    public String getEndPos() {
        return endPos;
    }

    public void setEndPos(String endPos) {
        this.endPos = endPos;
    }
}
