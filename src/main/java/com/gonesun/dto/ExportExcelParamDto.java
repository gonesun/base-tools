package com.gonesun.dto;

import lombok.Data;

import java.util.List;
import java.util.Map;

@Data
public class ExportExcelParamDto<T> extends DTO {
    private static final long serialVersionUID = 7903345837375699760L;

    public ExportExcelParamDto(){}

    public ExportExcelParamDto(Map<String, String> headMap, List<T> dataList, String fieldName, int sheetIndex){
        this.headMap = headMap;
        this.dataList = dataList;
        this.fieldName = fieldName;
        this.sheetIndex = sheetIndex;
    }

    public ExportExcelParamDto(List<Map<String, String>> headMapList, List<String> sheetNameList,
                               List<List<T>> batchDataList, String fieldName, int sheetIndex){
        this.headMapList = headMapList;
        this.sheetNameList = sheetNameList;
        this.batchDataList = batchDataList;
        this.fieldName = fieldName;
        this.sheetIndex = sheetIndex;
    }

    public ExportExcelParamDto(List<Map<String, String>> headMapList, List<String> sheetNameList,
                               List<List<T>> batchDataList, String fieldName, List<Integer> sheetIndexList){
        this.headMapList = headMapList;
        this.sheetNameList = sheetNameList;
        this.batchDataList = batchDataList;
        this.fieldName = fieldName;
        this.sheetIndexList = sheetIndexList;
    }

    /**
     * Excel模板路径
     */
    private String fieldName;

    /**
     * Excel中页签（0..）
     */
    private Integer sheetIndex;

    /**
     * Excel头数据
     */
    private Map<String, String> headMap;

    /**
     * 数据
     */
    private List<T> dataList;

    /**
     * 动态列
     */
    private List<String> unFixColList;

    /**
     * 动态列对应的名称字典
     */
    private Map<String, String> unFixColNameMap;

    /**
     * 页签名称
     */
    private String sheetName;

    /**
     * Excel中页签（0..）列表
     */
    private List<Integer> sheetIndexList;

    /**
     * Excel头数据
     */
    private List<Map<String, String>> headMapList;

    /**
     * 工作簿页签名称
     */
    private List<String> sheetNameList;

    /**
     * 数据
     */
    private List<List<T>> batchDataList;

    /**
     * 是否需要目录
     */
    private boolean needCatalog;

    /**
     * 是否不复制设置行行高
     * 0 复制
     * 1 不复制
     */
    private Integer notSetLineHeight;

    /**
     * 小数位控制字段
     */
    private List<String> decimalPlaceControlList;

    /**
     * 小数位控制字典
     * 不设置时按去尾零控制
     */
    private Map<String, Integer> decimalPlaceNumMap;

    /**
     * 动态列类型:(参照财务多栏账，生产入库成本测算打印)
     * #0、动态调整列（Excel模板中已经有这些列，只是需要调整顺序与列名，隐藏多余的动态列）
     * #1、动态添加列（Excel模板中没有这些列，需要动态添加。《注：该模式暂不支持合并单元格》）
     */
    private int unFixColsType = 0;

}
