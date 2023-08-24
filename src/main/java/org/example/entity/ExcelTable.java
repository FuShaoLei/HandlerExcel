package org.example.entity;

import java.util.List;

public class ExcelTable {

    /**
     * 设备编码
     */
    private String deviceCode;

    private List<FData> fDataList;

    public ExcelTable() {
    }

    public ExcelTable(String deviceCode) {
        this.deviceCode = deviceCode;
    }


    public List<FData> getStatusDataList() {
        return fDataList;
    }

    public void setStatusDataList(List<FData> fDataList) {
        this.fDataList = fDataList;
    }

    public String getDeviceCode() {
        return deviceCode;
    }



    public void setDeviceCode(String deviceCode) {
        this.deviceCode = deviceCode;
    }
}
