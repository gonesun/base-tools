package com.gonesun.tools;

import com.alibaba.fastjson.JSONObject;
import com.gonesun.dto.ExportExcelParamDto;
import com.gonesun.dto.ToolBusinessException;
import com.gonesun.utils.StringUtil;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

public class ExportExcelBaseService {
    private static final Logger logger = LoggerFactory.getLogger(ExportExcelBaseService.class);

    private final String emptyString = "";

    private final String specialTitleChar = ":/\\?*[]'：？／＼＊［］／＼";

    /**
     * 导出多页签
     * @param headMapList  Excel头数据
     * @param pageTitleList 页签名称列表
     * @param dataList 页签数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 报表Excel模板文件名（带路径的全名）
     * @param sheetIndex Excel中某个页签（0..）
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @return
     * @throws ToolBusinessException
     */
    public byte[] exportAll(List<Map<String, String>> headMapList, List<String> pageTitleList,
                                                   List<List<?>> dataList, String fieldName, int sheetIndex, List<String> unFixColList)
            throws ToolBusinessException {
        return exportAll(headMapList, pageTitleList, dataList, fieldName, sheetIndex, unFixColList, null);
    }

    /**
     * 导出多页签
     * @param headMapList  Excel头数据
     * @param pageTitleList 页签名称列表
     * @param dataList 页签数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 报表Excel模板文件名（带路径的全名）
     * @param sheetIndex Excel中某个页签（0..）
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @param userMap 自定义档案字典 <Code,Name>
     * @return
     * @throws ToolBusinessException
     */
    public byte[] exportAll(List<Map<String, String>> headMapList, List<String> pageTitleList,
                                                   List<List<?>> dataList, String fieldName, int sheetIndex, List<String> unFixColList,
                                                   Map<String, String> userMap) throws ToolBusinessException {
        return exportAll(headMapList, pageTitleList, dataList, fieldName, sheetIndex, unFixColList, userMap, false);
    }

    /**
     * 导出多页签
     * @param headMapList  Excel头数据
     * @param pageTitleList 页签名称列表
     * @param dataList 页签数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 报表Excel模板文件名（带路径的全名）
     * @param sheetIndex Excel中某个页签（0..）
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @param userMap 自定义档案字典 <Code,Name>
     * @param needCatalog 是否创建目录(第一个页签创建后面页签的超链接目录)
     * @return
     * @throws ToolBusinessException
     */
    public byte[] exportAll(List<Map<String, String>> headMapList, List<String> pageTitleList,
                                                   List<List<?>> dataList, String fieldName, int sheetIndex, List<String> unFixColList,
                                                   Map<String, String> userMap, boolean needCatalog) throws ToolBusinessException {
        ExportExcelParamDto paramDto = new ExportExcelParamDto(headMapList, pageTitleList, dataList, fieldName, sheetIndex);
        paramDto.setUnFixColList(unFixColList);
        paramDto.setUnFixColNameMap(userMap);
        paramDto.setNeedCatalog(needCatalog);
        return exportAll(paramDto);
    }

    /**
     * 导出多页签
     * @param headMapList  Excel头数据
     * @param pageTitleList 页签名称列表
     * @param dataList 页签数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 报表Excel模板文件名（带路径的全名）
     * @param sheetIndexList Excel中页签（0..）列表
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @param userMap 自定义档案字典 <Code,Name>
     * @return
     * @throws ToolBusinessException
     */
    public byte[] exportAll(List<Map<String, String>> headMapList, List<String> pageTitleList,
                                                   List<List<?>> dataList, String fieldName,
                                                   List<Integer> sheetIndexList, List<String> unFixColList,
                                                   Map<String, String> userMap) throws ToolBusinessException {
        return exportAll(headMapList, pageTitleList, dataList, fieldName, sheetIndexList, unFixColList, userMap, false);
    }

    /**
     * 导出多页签
     * @param headMapList  Excel头数据
     * @param pageTitleList 页签名称列表
     * @param dataList 页签数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 报表Excel模板文件名（带路径的全名）
     * @param sheetIndexList Excel中页签（0..）列表
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @param userMap 自定义档案字典 <Code,Name>
     * @param needCatalog 是否创建目录(第一个页签创建后面页签的超链接目录)
     * @return
     * @throws ToolBusinessException
     */
    public byte[] exportAll(List<Map<String, String>> headMapList, List<String> pageTitleList,
                                                   List<List<?>> dataList, String fieldName,
                                                   List<Integer> sheetIndexList, List<String> unFixColList,
                                                   Map<String, String> userMap, boolean needCatalog) throws ToolBusinessException {
        ExportExcelParamDto paramDto = new ExportExcelParamDto(headMapList, pageTitleList, dataList, fieldName, sheetIndexList);
        paramDto.setUnFixColList(unFixColList);
        paramDto.setUnFixColNameMap(userMap);
        paramDto.setNeedCatalog(needCatalog);
        return exportAll(paramDto);
    }

    public byte[] exportAll(ExportExcelParamDto paramDto) throws ToolBusinessException {
        List<Map<String, String>> headMapList = paramDto.getHeadMapList();
        List<String> pageTitleList = paramDto.getSheetNameList();
        List<List<?>> dataList = paramDto.getBatchDataList();
        String fieldName = paramDto.getFieldName();
        List<Integer> sheetIndexList = paramDto.getSheetIndexList();
        Integer sheetIndex = paramDto.getSheetIndex();
        List<String> unFixColList = paramDto.getUnFixColList();
        Map<String, String> userMap = paramDto.getUnFixColNameMap();
        boolean needCatalog = paramDto.isNeedCatalog();
        int allCount = dataList.size();
        if (fieldName == null) {
            throw new ToolBusinessException("80564", "报表文件名称不能为空！");
        }
        if (dataList.isEmpty()) {
            throw new ToolBusinessException("80564", "没有要导出的数据！");
        }
        // 是否多模板
        boolean isDiffIndex = false;
        if(null != sheetIndexList) {
            isDiffIndex = true;
            if (dataList.size() != headMapList.size() || dataList.size() != pageTitleList.size()
                    || dataList.size() != sheetIndexList.size()) {
                throw new ToolBusinessException("80564", "传入数据不一致！");
            }
        }else{
            if (null == sheetIndex || sheetIndex < 0) {
                sheetIndex = 0;
            }
        }
        ExcelBaseInfoDto excelBaseInfoDto = new ExcelBaseInfoDto();
        excelBaseInfoDto.dataList = dataList.get(0);

        excelBaseInfoDto.fieldFullName = fieldName;
        try {
            if(isDiffIndex){
                this.init(unFixColList, sheetIndexList, excelBaseInfoDto, userMap);
            }else {
                this.init(unFixColList, sheetIndexList, excelBaseInfoDto, userMap);
            }
        } catch (Exception ex) {
            throw new ToolBusinessException("80562", "导出Excel失败【初始化表格失败】！", null, ex);
        }
        excelBaseInfoDto.decimalPlaceControlList = paramDto.getDecimalPlaceControlList();
        excelBaseInfoDto.decimalPlaceNumMap = paramDto.getDecimalPlaceNumMap();
        int modelSheetCount = excelBaseInfoDto.workbook.getNumberOfSheets();
        Map<String, Integer> titleMap = new HashMap<>();
        String catalogTitle = "目 录";
        if(needCatalog) {
            titleMap.put(catalogTitle, 1);
        }
        List<String> titleList = new ArrayList<>();
        int sheetIndexTmp;
        for (int i = 0; i < allCount; i++) {
            if(isDiffIndex){
                sheetIndexTmp = sheetIndexList.get(i);
                excelBaseInfoDto.setConfigBySon(sheetIndexTmp);
            }else{
                sheetIndexTmp = sheetIndex;
            }
            excelBaseInfoDto.workbook.cloneSheet(sheetIndexTmp);
            excelBaseInfoDto.sheet = excelBaseInfoDto.workbook.getSheetAt(modelSheetCount + i);
            commitHead(headMapList.get(i), excelBaseInfoDto);
            excelBaseInfoDto.dataList = dataList.get(i);
            excelBaseInfoDto.clearColMergeMap();
            String pageTitle = pageTitleList.get(i);
            char[] ac = pageTitle.toCharArray();
            StringBuilder newTitle = new StringBuilder();
            // 特殊字符屏蔽
            for (char c : ac) {
                if (specialTitleChar.indexOf(c) < 0) {
                    newTitle.append(c);
                }else {
                    newTitle.append(" ");
                }
            }
            boolean isSub = false;
            if(newTitle.length() > 31){
                newTitle.delete(30, newTitle.length());
                isSub = true;
            }
            // 增加title重复判断
            Integer count = titleMap.get(newTitle.toString());
            // 页签名称超长截取后，有重复的，再次多截取，再有重复的就增加序号，序号再超了就暂时不管了
            if(null != count && isSub){
                newTitle.delete(25, newTitle.length());
                count = titleMap.get(newTitle.toString());
            }
            while (null != count){
                titleMap.put(newTitle.toString(), count + 1);
                newTitle.append(count + 1);
                count = titleMap.get(newTitle.toString());
            }
            titleMap.put(newTitle.toString(), 1);
            titleList.add(newTitle.toString());
            excelBaseInfoDto.workbook.setSheetName(modelSheetCount + i, newTitle.toString());
            try {
                for (int j = 0; j < excelBaseInfoDto.dataList.size(); j++) {
                    setRowFromObject(j, excelBaseInfoDto);
                }
            } catch (Exception ex) {
                throw new ToolBusinessException("80562", "导出Excel失败【写入数据失败】！", null, ex);
            }
        }
        // 删除所有模板
        while (modelSheetCount > 0) {
            excelBaseInfoDto.workbook.removeSheetAt(0);
            modelSheetCount-- ;
        }
        if(needCatalog) {
            Sheet sheet = excelBaseInfoDto.workbook.createSheet(catalogTitle);
            excelBaseInfoDto.workbook.setSheetOrder(catalogTitle, 0);
            createCatalog(excelBaseInfoDto.workbook, sheet, titleList);
        }
        byte[] returnByte = null;

        try {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            excelBaseInfoDto.workbook.write(byteArrayOutputStream);
            byteArrayOutputStream.flush();
            byteArrayOutputStream.close();
            returnByte = byteArrayOutputStream.toByteArray();
        } catch (Exception ex) {
            throw new ToolBusinessException("80562", "导出Excel失败【转换数据失败】", null, ex);
        }
        return returnByte;
    }

    /**
     * 导出Excel文件
     * @param headMap Excel头数据
     * @param dataList 数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 全名，包含路径及其文件扩展名.xls(例如/template/gl/balanceSumAuxRpt.xls),请按不同的领域创建例如/template/xx/xxx.xls
     * @param sheetIndex Excel中某个页签（0..）
     * @return
     * @throws ToolBusinessException
     */
    public byte[] export(Map<String, String> headMap, List<?> dataList, String fieldName,
                                                int sheetIndex) throws ToolBusinessException {
        return export(headMap, dataList, fieldName, sheetIndex, null);
    }

    /**
     * 导出Excel文件
     * @param headMap Excel头数据
     * @param dataList 数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 全名，包含路径及其文件扩展名.xls(例如/template/gl/balanceSumAuxRpt.xls),请按不同的领域创建例如/template/xx/xxx.xls
     * @param sheetIndex Excel中某个页签（0..）
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @return
     * @throws ToolBusinessException
     */
    public byte[] export(Map<String, String> headMap, List<?> dataList, String fieldName,
                                                int sheetIndex, List<String> unFixColList) throws ToolBusinessException {
        return export(headMap, dataList, fieldName, sheetIndex, unFixColList, null);
    }

    /**
     * 导出Excel文件
     * @param headMap Excel头数据
     * @param dataList 数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 全名，包含路径及其文件扩展名.xls(例如/template/gl/balanceSumAuxRpt.xls),请按不同的领域创建例如/template/xx/xxx.xls
     * @param sheetIndex Excel中某个页签（0..）
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @param userMap 自定义档案字典 <Code,Name>
     * @return
     * @throws ToolBusinessException
     */
    public byte[] export(Map<String, String> headMap, List<?> dataList, String fieldName,
                                                int sheetIndex, List<String> unFixColList, Map<String, String> userMap) throws ToolBusinessException {
        return export(headMap, dataList, fieldName, sheetIndex, unFixColList, userMap, null);
    }

    /**
     * 导出Excel文件
     * @param headMap Excel头数据
     * @param dataList 数据列表 com.ttk.edf.base.DTO 的子类
     * @param fieldName 全名，包含路径及其文件扩展名.xls(例如/template/gl/balanceSumAuxRpt.xls),请按不同的领域创建例如/template/xx/xxx.xls
     * @param sheetIndex Excel中某个页签（0..）
     * @param unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @param userMap 自定义档案字典 <Code,Name>
     * @return
     * @throws ToolBusinessException
     */
    public byte[] export(Map<String, String> headMap, List<?> dataList, String fieldName,
                                                int sheetIndex, List<String> unFixColList, Map<String, String> userMap,
                                                String sheetName, Integer... notSetLineHeight) throws ToolBusinessException {
        ExportExcelParamDto paramDto = new ExportExcelParamDto(headMap, dataList, fieldName, sheetIndex);
        paramDto.setUnFixColList(unFixColList);
        paramDto.setUnFixColNameMap(userMap);
        paramDto.setSheetName(sheetName);
        if (notSetLineHeight.length != 0) {
            paramDto.setNotSetLineHeight(notSetLineHeight[0]);
        }
        return export(paramDto);
    }

    /**
     * 导出Excel文件 （首次使用 生产入库，不定数量动态列）
     * @param headMap
     * @param dataList
     * @param fieldName
     * @param sheetIndex
     * @param unFixColList
     * @param userMap
     * @return
     * @throws ToolBusinessException
     */
    public byte[] export(Map<String, String> headMap, List<?> dataList, String fieldName,
                                                int sheetIndex, Map<String, String> userMap, List<String> unFixColList ) throws ToolBusinessException {
        ExportExcelParamDto paramDto = new ExportExcelParamDto(headMap, dataList, fieldName, sheetIndex);
        paramDto.setUnFixColList(unFixColList);
        paramDto.setUnFixColNameMap(userMap);
        paramDto.setUnFixColsType(1);
        return export(paramDto);
    }

    /**
     * 导出Excel文件
     * @param paramDto#headMap Excel头数据
     * @see ExportExcelParamDto#dataList 数据列表 com.ttk.edf.base.DTO 的子类
     * @see ExportExcelParamDto#fieldName 全名，包含路径及其文件扩展名.xls(例如/template/gl/balanceSumAuxRpt.xls),请按不同的领域创建例如/template/xx/xxx.xls
     * @see ExportExcelParamDto#sheetIndex Excel中某个页签（0..）
     * @see ExportExcelParamDto#unFixColList Excel变动列数据（变动列只支持在模板最前面）
     * @see ExportExcelParamDto#unFixColNameMap 自定义档案字典 <Code,Name>
     * @return
     * @throws ToolBusinessException
     */
    public byte[] export(ExportExcelParamDto paramDto) throws ToolBusinessException {
        if (paramDto.getFieldName() == null) {
            throw new ToolBusinessException("80564", "报表文件名称不能为空！");
        }
        Integer sheetIndex = paramDto.getSheetIndex();
        if (sheetIndex == null || sheetIndex < 0) {
            sheetIndex = 0;
        }
        ExcelBaseInfoDto excelBaseInfoDto = new ExcelBaseInfoDto();
        excelBaseInfoDto.dataList = paramDto.getDataList();
        excelBaseInfoDto.fieldFullName = paramDto.getFieldName();
        excelBaseInfoDto.notSetLineHeight = paramDto.getNotSetLineHeight() == null ? 0 : paramDto.getNotSetLineHeight();
        excelBaseInfoDto.decimalPlaceControlList = paramDto.getDecimalPlaceControlList();
        excelBaseInfoDto.decimalPlaceNumMap = paramDto.getDecimalPlaceNumMap();
        byte[] bytes = null;
        try {
            if(paramDto.getUnFixColsType() == 0) {
                this.init(paramDto.getUnFixColList(), sheetIndex, excelBaseInfoDto, paramDto.getUnFixColNameMap());
            }else{
                this.initWithUnFix(paramDto.getUnFixColList(), sheetIndex, excelBaseInfoDto, paramDto.getUnFixColNameMap());
            }
            commitHead(paramDto.getHeadMap(), excelBaseInfoDto);
            for (int i = 0; i < excelBaseInfoDto.dataList.size(); i++) {
                setRowFromObject(i, excelBaseInfoDto);
            }
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            excelBaseInfoDto.workbook.write(byteArrayOutputStream);
            if (!StringUtil.isNullOrEmpty(paramDto.getSheetName())) {
                excelBaseInfoDto.workbook.setSheetName(0, paramDto.getSheetName());
            }
            byteArrayOutputStream.flush();
            byteArrayOutputStream.close();
            bytes = byteArrayOutputStream.toByteArray();
        } catch (Exception ex) {
            throw new ToolBusinessException("80562", "导出Excel失败！", null, ex);
        }
        return bytes;
    }

    private Workbook loadFile(InputStream inputStream, String fieldFullName){
        Workbook workbook;
        try {
            boolean isXls = fieldFullName.toLowerCase().endsWith(".xls");
            boolean needClose = false;
            if (inputStream == null) {
                inputStream = this.getClass().getResourceAsStream(fieldFullName);
                needClose = true;
            }
            if(isXls) {
                workbook = new HSSFWorkbook(inputStream);
            }else{
                workbook = new XSSFWorkbook(inputStream);
            }
            if(needClose){
                inputStream.close();
            }
            return workbook;
        } catch (IOException ex) {
            throw new ToolBusinessException("80562", "读取Excel模板文件异常！", null, ex);
        }
    }

    /**
     * 创建超链接目录
     * @param wb
     * @param sheet
     * @param list
     */
    private void createCatalog(Workbook wb, Sheet sheet, List<String> list) {
        sheet.setColumnWidth(0, 300);
        sheet.setColumnWidth(1, sheet.getColumnWidth(1) * 3);
        for (int i = 0; i < list.size(); i++) {
            String title = list.get(i);
            Row row = sheet.createRow(i);

            /* 连接跳转*/
            Cell likeCell = row.createCell((short)1);
//			setCellFormula
            Hyperlink hyperlink = wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
            // "#"表示本文档    "明细页面"表示sheet页名称  "A1"表示第几列第几行
            hyperlink.setAddress(String.format("'%s'!A1", title));
            likeCell.setHyperlink(hyperlink);
            // 点击进行跳转
            likeCell.setCellValue(title);

            /* 设置为超链接的样式*/
            CellStyle linkStyle = wb.createCellStyle();
            Font cellFont= wb.createFont();
            cellFont.setUnderline((byte) 1);
            cellFont.setColor(IndexedColors.BLUE.getIndex());
            linkStyle.setFont(cellFont);
            likeCell.setCellStyle(linkStyle);
        }
    }

    // 构造表头可变信息
    private void commitHead(Map<String, String> headMap,
                                                   ExcelBaseInfoDto excelBaseInfoDto) {
        for (FileDataInfo headInfo : excelBaseInfoDto.headAndEbdConfig) {
            Row row = excelBaseInfoDto.sheet.getRow(headInfo.rowIndex);
            if (null == row) {
                continue;
            }
            Cell cell = row.getCell(headInfo.cellIndex);
            String cellValue = cell.getStringCellValue();
            if (null == cellValue) {
                continue;
            }
            String replaceValue = headMap.get(headInfo.fieldName);
            if (null != replaceValue) {
                cellValue = cellValue.replace(String.format("{%s}", headInfo.fieldName), replaceValue);
            }
            cell.setCellValue(cellValue);
        }
    }

    // 初始化读取Excel模板数据
    private void init(List<String> unFixColList, int sheetIndex,
                                             ExcelBaseInfoDto excelBaseInfoDto, Map<String, String> userMap) throws Exception {
        // 读取Excel模板数据
        excelBaseInfoDto.workbook = loadFile(excelBaseInfoDto.inputStream, excelBaseInfoDto.fieldFullName);

        try {
            for (int i = excelBaseInfoDto.workbook.getNumberOfSheets() - 1; i >= 0; i--) {
                // 删除Sheet
                if (i == sheetIndex) {
                    continue;
                }
                excelBaseInfoDto.workbook.removeSheetAt(i);
            }
            sheetIndex = 0;
        } catch (Exception ex) {
            throw new ToolBusinessException("80562", "解析Excel模板页签异常！", null, ex);
        }

        initModelDetail(unFixColList, sheetIndex, excelBaseInfoDto, userMap);
    }

    // 初始化读取Excel模板数据
    private void init(List<String> unFixColList, List<Integer> sheetIndexList,
                                             ExcelBaseInfoDto excelBaseInfoDto, Map<String, String> userMap) throws Exception {
        // 读取Excel模板数据
        excelBaseInfoDto.workbook = loadFile(excelBaseInfoDto.inputStream, excelBaseInfoDto.fieldFullName);

        List<Integer> handledSheetIndexList = new ArrayList<>();
        for (Integer curSheetIndex : sheetIndexList) {
            if (handledSheetIndexList.contains(curSheetIndex)) {
                continue;
            }
            handledSheetIndexList.add(curSheetIndex);
            initSon(unFixColList, curSheetIndex, excelBaseInfoDto, userMap);
        }
    }

    // 初始化读取Excel模板数据
    private void initSon(List<String> unFixColList, int sheetIndex,
                                                ExcelBaseInfoDto excelBaseInfoDto, Map<String, String> userMap) throws Exception {
        excelBaseInfoDto.clearConfig();
        sheetIndex = initModelDetail(unFixColList, sheetIndex, excelBaseInfoDto, userMap);

        ExcelBaseInfoDto sonDto = excelBaseInfoDto.clone(sheetIndex);
        if (excelBaseInfoDto.sonList == null) {
            excelBaseInfoDto.sonList = new ArrayList<>();
        }
        excelBaseInfoDto.sonList.add(sonDto);
    }

    private int initModelDetail(List<String> unFixColList, int sheetIndex, ExcelBaseInfoDto excelBaseInfoDto, Map<String, String> userMap) {
        if (sheetIndex >= excelBaseInfoDto.workbook.getNumberOfSheets()) {
            sheetIndex = excelBaseInfoDto.workbook.getNumberOfSheets() - 1;
        }
        excelBaseInfoDto.sheet = excelBaseInfoDto.workbook.getSheetAt(sheetIndex);
        // 读取模板配置行信息
        excelBaseInfoDto.headRowCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(0).getNumericCellValue();
        excelBaseInfoDto.colCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(1).getNumericCellValue();
        excelBaseInfoDto.headConfigCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(2).getNumericCellValue();
        excelBaseInfoDto.rowHeight = excelBaseInfoDto.sheet.getRow(1).getHeight();
        // 读取动态列头配置数据
        if (null != unFixColList && !unFixColList.isEmpty()) {
            try {
                excelBaseInfoDto.unFixColRowNo = (int) excelBaseInfoDto.sheet.getRow(0).getCell(3)
                        .getNumericCellValue();
                excelBaseInfoDto.headUnFixColCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(4)
                        .getNumericCellValue();
            } catch (Exception ex) {
                excelBaseInfoDto.unFixColRowNo = 0;
                excelBaseInfoDto.headUnFixColCount = 0;
            }
        } else {
            excelBaseInfoDto.unFixColRowNo = 0;
            excelBaseInfoDto.headUnFixColCount = 0;
        }
        try {
            excelBaseInfoDto.endRowCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(5)
                    .getNumericCellValue();
            excelBaseInfoDto.endConfigCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(6)
                    .getNumericCellValue();
        } catch (Exception ex) {
            excelBaseInfoDto.endRowCount = 0;
            excelBaseInfoDto.endConfigCount = 0;
        }
        // 动态列可用，则调整动态列，隐藏多余的动态列
        if (excelBaseInfoDto.unFixColRowNo > 0 && excelBaseInfoDto.headUnFixColCount > 0) {
            int findCount = 0;
            for (int j = 0; j < unFixColList.size() && j < excelBaseInfoDto.headUnFixColCount; j++) {
                String fieldName = unFixColList.get(j);
                boolean isFind = false;
                for (int i = 0; i < excelBaseInfoDto.headUnFixColCount; i++) {
                    Cell cell = excelBaseInfoDto.sheet.getRow(1).getCell(i);
                    if (fieldName.equals(cell.getStringCellValue())) {
                        String newName = null;
                        if (userMap != null) {
                            newName = userMap.get(fieldName);
                        }
                        //交换列宽
                        int width_i = excelBaseInfoDto.sheet.getColumnWidth(i);
                        int width_j = excelBaseInfoDto.sheet.getColumnWidth(j);
                        excelBaseInfoDto.sheet.setColumnWidth(i, width_j);
                        excelBaseInfoDto.sheet.setColumnWidth(j, width_i);
                        // 动态列的列头互换
                        Cell cellA = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(j);
                        Cell cellB = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(i);
                        if (!StringUtil.isEmpty(newName)) {
                            cellB.setCellValue(newName);
                        }
                        exchangeCell(cellA, cellB);
                        //交换配置头信息
                        FileDataInfo fileDataInfo = new FileDataInfo(j, fieldName);
                        fileDataInfo.cellStyle = cell.getCellStyle();
                        excelBaseInfoDto.bodyConfig.add(fileDataInfo);
                        managerConfigInit(i, fileDataInfo, excelBaseInfoDto);
                        cellA = excelBaseInfoDto.sheet.getRow(1).getCell(j);
                        cellB = excelBaseInfoDto.sheet.getRow(1).getCell(i);
                        exchangeCell(cellA, cellB);
                        isFind = true;
                        findCount++;
                        break;
                    }
                }
                if (!isFind) {
                    logger.error(String.format("动态列【%s】在模板中未找到对应列", fieldName));
                }
            }
            // 隐藏多余的动态列
            for (int j = findCount; j < excelBaseInfoDto.headUnFixColCount; j++) {
                Cell cell = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(j);
                cell.setCellValue("");
                excelBaseInfoDto.sheet.setColumnWidth(j, 0);
                excelBaseInfoDto.needMergedRegion = true;
                excelBaseInfoDto.firstColumn = findCount - 1;
                excelBaseInfoDto.lastColumn = excelBaseInfoDto.headUnFixColCount - 1;
                FileDataInfo fileDataInfo = new FileDataInfo(j, "");
                fileDataInfo.cellStyle = excelBaseInfoDto.bodyConfig.get(excelBaseInfoDto.bodyConfig.size() - 1).cellStyle;
                excelBaseInfoDto.bodyConfig.add(fileDataInfo);
            }
            if (excelBaseInfoDto.needMergedRegion) {
                try {
                    int sheetMergeCount = excelBaseInfoDto.sheet.getNumMergedRegions();
                    for (int i = 0; i < sheetMergeCount; i++) {
                        CellRangeAddress range = excelBaseInfoDto.sheet.getMergedRegion(i);
                        if (range == null) {
                            continue;
                        }
                        int firstColumn = range.getFirstColumn();
                        int lastColumn = range.getLastColumn();
                        int firstRow = range.getFirstRow();
                        int lastRow = range.getLastRow();
                        if (firstRow == excelBaseInfoDto.unFixColRowNo - 1 && firstColumn >= findCount - 1 && firstColumn <= excelBaseInfoDto.headUnFixColCount - 1) {
                            excelBaseInfoDto.sheet.removeMergedRegion(i);
                            i--;
                        }
                    }
                    int firstRow = excelBaseInfoDto.unFixColRowNo - 1;
                    int lastRow = excelBaseInfoDto.headRowCount + excelBaseInfoDto.configureRowCount - 1;
                    excelBaseInfoDto.sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, excelBaseInfoDto.firstColumn, excelBaseInfoDto.lastColumn));
                } catch (Exception ex) {

                }
            }
        } else {
            excelBaseInfoDto.unFixColRowNo = 0;
            excelBaseInfoDto.headUnFixColCount = 0;
        }
        for (int i = excelBaseInfoDto.headUnFixColCount; i < excelBaseInfoDto.colCount; i++) {
            Cell cell = excelBaseInfoDto.sheet.getRow(1).getCell(i);
            FileDataInfo fileDataInfo = new FileDataInfo(i, cell.getStringCellValue());
            fileDataInfo.cellStyle = cell.getCellStyle();
            excelBaseInfoDto.bodyConfig.add(fileDataInfo);
            managerConfigInit(i, fileDataInfo, excelBaseInfoDto);
        }
        for (int i = 0; i < excelBaseInfoDto.headConfigCount + excelBaseInfoDto.endConfigCount; i++) {
            Cell cell = excelBaseInfoDto.sheet.getRow(2).getCell(i);
            if (null == cell) {
                continue;
            }
            String tmp = cell.getStringCellValue();
            String[] heads = tmp.split(":");
            if (heads.length != 3) {
                continue;
            }
            FileDataInfo fileDataInfo = new FileDataInfo(
                    Integer.valueOf(heads[0]) - 1 - excelBaseInfoDto.configureRowCount, Integer.valueOf(heads[1]) - 1,
                    heads[2]);
            excelBaseInfoDto.headAndEbdConfig.add(fileDataInfo);
        }
        for (int i = excelBaseInfoDto.configureRowCount; i < excelBaseInfoDto.headRowCount
                + excelBaseInfoDto.configureRowCount + excelBaseInfoDto.endRowCount; i++) {
            excelBaseInfoDto.headAndEndRowHeightList.add(excelBaseInfoDto.sheet.getRow(i).getHeight());
        }
        // 删除配置行
        // 3.16poi shiftRows 不会带合并单元格信息，手工复制过去
        List<CellRangeAddress> list = new ArrayList<>();
        for (int j = excelBaseInfoDto.configureRowCount; j <= excelBaseInfoDto.headRowCount
                + excelBaseInfoDto.endRowCount + excelBaseInfoDto.configureRowCount; j++) {
            for (int i = 0; i < excelBaseInfoDto.sheet.getNumMergedRegions(); i++) {
                CellRangeAddress cellRangeAddress = excelBaseInfoDto.sheet.getMergedRegion(i);
                if (cellRangeAddress.getFirstRow() == j) {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(j - excelBaseInfoDto.configureRowCount,
                            (j - excelBaseInfoDto.configureRowCount
                                    + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                            cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                    list.add(newCellRangeAddress);
                }
            }
        }
        for (int i = 0; i < excelBaseInfoDto.configureRowCount; i++) {
            excelBaseInfoDto.sheet.removeRow(excelBaseInfoDto.sheet.getRow(i));
        }
        excelBaseInfoDto.sheet.shiftRows(excelBaseInfoDto.configureRowCount, excelBaseInfoDto.sheet.getLastRowNum(),
                -excelBaseInfoDto.configureRowCount, true, true);

        for (CellRangeAddress address : list) {
            excelBaseInfoDto.sheet.addMergedRegionUnsafe(address);
        }

        for (int i = 0; i < excelBaseInfoDto.headRowCount; i++) {
            excelBaseInfoDto.sheet.getRow(i).setHeight(excelBaseInfoDto.headAndEndRowHeightList.get(i));
        }
        // 初始化字段类型数据
        initFieldType(excelBaseInfoDto);
        return sheetIndex;
    }

    /**
     * 合并配置列初始化
     * @param colNo 列号
     * @param fileDataInfo 当前单元格配置信息
     * @param excelBaseInfoDto Excel 配置对象
     */
    private void managerConfigInit(int colNo, FileDataInfo fileDataInfo, ExcelBaseInfoDto excelBaseInfoDto){

        String fieldName = fileDataInfo.fieldName;
        // 支持两种格式
        // 1.accountName:mergeKey:accountName#voucherDate  列名:合并列标识:合并列分组条件（多列组合分组#号分割），严格按照配置的分组条件分组
        // 2.accountName:mergeKey 列名:合并列标识，作为所有列的合并分组标识，只能有一列配置
        if (fieldName.contains(":")) {
            String[] fieldNames = fieldName.split(":");
            fieldName = fieldNames[0];
            fileDataInfo.fieldName = fieldName;
            String attr = fieldNames[1];
            if (attr.equals("merge")) {
                fileDataInfo.mergeRow = true;
            } else if (attr.equals("mergeKey")) {
                fileDataInfo.mergeRow = true;
                excelBaseInfoDto.mergeKeyCol = fileDataInfo;
                // 多列
                if(fieldNames.length == 3){
                    String keyStr = fieldNames[2];
                    String[] keyArr = keyStr.split("#");
                    fileDataInfo.mergeFieldNameList.addAll(Arrays.asList(keyArr));
                    excelBaseInfoDto.mergeKeyCol = null;
                }
            }
        }
        if (fileDataInfo.mergeRow) {
            excelBaseInfoDto.colMergeMap.put(colNo, new HashMap<>());
        }
    }

    /***
     * 多于两个的不定动态列表格处理
     * @param unFixColList
     * @param sheetIndex
     * @param excelBaseInfoDto
     * @param userMap
     * @return
     */
    private int initWithUnFix(List<String> unFixColList, int sheetIndex, ExcelBaseInfoDto excelBaseInfoDto, Map<String, String> userMap) {
        // 读取Excel模板数据
        excelBaseInfoDto.workbook = loadFile(excelBaseInfoDto.inputStream, excelBaseInfoDto.fieldFullName);

        try {
            for (int i = excelBaseInfoDto.workbook.getNumberOfSheets() - 1; i >= 0; i--) {
                // 删除Sheet
                if (i == sheetIndex) {
                    continue;
                }
                excelBaseInfoDto.workbook.removeSheetAt(i);
            }
            sheetIndex = 0;
        } catch (Exception ex) {
            throw new ToolBusinessException("80562", "解析Excel模板页签异常！", null, ex);
        }
        if (sheetIndex >= excelBaseInfoDto.workbook.getNumberOfSheets()) {
            sheetIndex = excelBaseInfoDto.workbook.getNumberOfSheets() - 1;
        }
        excelBaseInfoDto.sheet = excelBaseInfoDto.workbook.getSheetAt(sheetIndex);
        // 读取模板配置行信息
        excelBaseInfoDto.headRowCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(0).getNumericCellValue();
        excelBaseInfoDto.colCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(1).getNumericCellValue();
        excelBaseInfoDto.headConfigCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(2).getNumericCellValue();
        excelBaseInfoDto.rowHeight = excelBaseInfoDto.sheet.getRow(1).getHeight();
        if(null == unFixColList){
            unFixColList = new ArrayList<>();
        }
        // 读取动态列头配置数据
        if (!unFixColList.isEmpty()) {
            try {
                excelBaseInfoDto.unFixColRowNo = (int) excelBaseInfoDto.sheet.getRow(0).getCell(3)
                        .getNumericCellValue();
                excelBaseInfoDto.headUnFixColCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(4)
                        .getNumericCellValue();
            } catch (Exception ex) {
                excelBaseInfoDto.unFixColRowNo = 0;
                excelBaseInfoDto.headUnFixColCount = 0;
            }
        } else {
            excelBaseInfoDto.unFixColRowNo = 0;
            excelBaseInfoDto.headUnFixColCount = 0;
        }
        try {
            excelBaseInfoDto.endRowCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(5)
                    .getNumericCellValue();
            excelBaseInfoDto.endConfigCount = (int) excelBaseInfoDto.sheet.getRow(0).getCell(6)
                    .getNumericCellValue();
        } catch (Exception ex) {
            excelBaseInfoDto.endRowCount = 0;
            excelBaseInfoDto.endConfigCount = 0;
        }

        for (int i = 0; i < excelBaseInfoDto.headConfigCount + excelBaseInfoDto.endConfigCount; i++) {
            Cell cell = excelBaseInfoDto.sheet.getRow(2).getCell(i);
            if (null == cell) {
                continue;
            }
            String tmp = cell.getStringCellValue();
            String[] heads = tmp.split(":");
            if (heads.length != 3) {
                continue;
            }
            FileDataInfo fileDataInfo = new FileDataInfo(
                    Integer.valueOf(heads[0]) - 1 - excelBaseInfoDto.configureRowCount, Integer.valueOf(heads[1]) - 1,
                    heads[2]);
            excelBaseInfoDto.headAndEbdConfig.add(fileDataInfo);
        }
        for (int i = excelBaseInfoDto.configureRowCount; i < excelBaseInfoDto.headRowCount
                + excelBaseInfoDto.configureRowCount + excelBaseInfoDto.endRowCount; i++) {
            excelBaseInfoDto.headAndEndRowHeightList.add(excelBaseInfoDto.sheet.getRow(i).getHeight());
        }
        // 删除配置行
        // 3.16poi shiftRows 不会带合并单元格信息，手工复制过去
        List<CellRangeAddress> list = new ArrayList<>();
        for (int j = excelBaseInfoDto.configureRowCount; j <= excelBaseInfoDto.headRowCount
                + excelBaseInfoDto.endRowCount + excelBaseInfoDto.configureRowCount; j++) {
            for (int i = 0; i < excelBaseInfoDto.sheet.getNumMergedRegions(); i++) {
                CellRangeAddress cellRangeAddress = excelBaseInfoDto.sheet.getMergedRegion(i);
                if (cellRangeAddress.getFirstRow() == j) {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(j - excelBaseInfoDto.configureRowCount,
                            (j - excelBaseInfoDto.configureRowCount
                                    + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                            cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                    list.add(newCellRangeAddress);
                }
            }
        }
        // 插入单元格
        int allCellCount = excelBaseInfoDto.colCount;
        int rowCount = excelBaseInfoDto.headRowCount	+ excelBaseInfoDto.endRowCount;
        int insertCount = 0;
        int insertStart = 0;
        insertCount = unFixColList.size() - 2;
        insertStart = excelBaseInfoDto.headUnFixColCount;
        if(insertCount > 0) {
            excelBaseInfoDto.colCount = allCellCount + insertCount;
        }else{
            insertCount = 0;
        }
        for(int j = 0; j < excelBaseInfoDto.configureRowCount + rowCount; j++){
            for (int i = 0; i < list.size(); i++) {
                excelBaseInfoDto.sheet.removeMergedRegion(i);
            }
        }
        // 添加动态列并设置显示名称
        if(unFixColList.size() == 1){
            Cell cell1 = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(insertStart - 1);
            cell1.setCellValue(userMap.get(unFixColList.get(0)));
            Cell cell1_ = excelBaseInfoDto.sheet.getRow(1).getCell(insertStart - 1);
            cell1_.setCellValue(unFixColList.get(0));
        } else if(unFixColList.size() == 2) {
            Cell cell1 = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(insertStart - 1);
            cell1.setCellValue(userMap.get(unFixColList.get(0)));
            Cell cell1_ = excelBaseInfoDto.sheet.getRow(1).getCell(insertStart - 1);
            cell1_.setCellValue(unFixColList.get(0));
            Cell cell2 = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(insertStart);
            cell2.setCellValue(userMap.get(unFixColList.get(1)));
            Cell cell2_ = excelBaseInfoDto.sheet.getRow(1).getCell(insertStart);
            cell2_.setCellValue(unFixColList.get(1));
        }else if(unFixColList.size() > 2){
            for (int i = 1; i < excelBaseInfoDto.configureRowCount + rowCount; i++) {
                if(i != 1 && i < excelBaseInfoDto.configureRowCount){
                    continue;
                }
                Row row = excelBaseInfoDto.sheet.getRow(i);
                for (int k = allCellCount + insertCount - 1; k >= insertStart; k--) {
                    Cell cell;
                    if (k >= allCellCount) {
                        cell = row.createCell(k);
                    } else {
                        cell = row.getCell(k);
                        if(null == cell){
                            cell = row.createCell(k);
                        }
                    }
                    if (k >= insertStart + insertCount) {
                        if (i == excelBaseInfoDto.configureRowCount) {
                            excelBaseInfoDto.sheet.setColumnWidth(k, excelBaseInfoDto.sheet.getColumnWidth(k - insertCount));
                        }
                        Cell fCell = row.getCell(k - insertCount);
                        if (null != fCell) {
                            cell.setCellStyle(fCell.getCellStyle());
                            cell.setCellValue(fCell.getStringCellValue());

                        }
                    } else {
                        if (i == excelBaseInfoDto.configureRowCount) {
                            excelBaseInfoDto.sheet.setColumnWidth(k, excelBaseInfoDto.sheet.getColumnWidth(insertStart));
                        }
                        Cell fCell = row.getCell(insertStart);
                        if (null != fCell) {
                            cell.setCellStyle(fCell.getCellStyle());
                            cell.setCellValue(fCell.getStringCellValue());
                        }
                        cell.setCellValue("");
                    }
                    if(i == excelBaseInfoDto.unFixColRowNo - 1 && k < insertStart + insertCount){
                        CellStyle style = cell.getCellStyle();
                        style.setWrapText(true);
                        cell.setCellStyle(style);
                        ;
                    }
                }
            }
            for (int k = insertStart; k < insertStart + insertCount + 2; k++) {
                Cell cell = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).getCell(k - 1);
                if(null == cell){
                    cell = excelBaseInfoDto.sheet.getRow(excelBaseInfoDto.unFixColRowNo - 1).createCell(k - 1);
                }
                cell.setCellValue(userMap.get(unFixColList.get(k - insertStart)));
                Cell cell_ = excelBaseInfoDto.sheet.getRow(1).getCell(k - 1);
                if(null == cell_){
                    cell_ = excelBaseInfoDto.sheet.getRow(1).createCell(k - 1);
                }
                cell_.setCellValue(unFixColList.get(k - insertStart));
            }
        }
        // 整理列头
        for(int i = 0; i < excelBaseInfoDto.colCount; i++){
            Cell cell = excelBaseInfoDto.sheet.getRow(1).getCell(i);
            FileDataInfo fileDataInfo = new FileDataInfo(cell.getColumnIndex(), cell.getStringCellValue());
            fileDataInfo.cellStyle = cell.getCellStyle();
            excelBaseInfoDto.bodyConfig.add(fileDataInfo);
            managerConfigInit(i, fileDataInfo, excelBaseInfoDto);
        }
        // 移除配置行
        for (int i = 0; i < excelBaseInfoDto.configureRowCount; i++) {
            excelBaseInfoDto.sheet.removeRow(excelBaseInfoDto.sheet.getRow(i));
        }
        excelBaseInfoDto.sheet.shiftRows(excelBaseInfoDto.configureRowCount, excelBaseInfoDto.sheet.getLastRowNum(),
                -excelBaseInfoDto.configureRowCount, true, true);
        // 设置单元格合并
        for (CellRangeAddress address : list) {
            int firstColumn = address.getFirstColumn();
            int lastColumn = address.getLastColumn();
            if(firstColumn >= insertStart - 1){
                if(firstColumn != insertStart - 1) {
                    address.setFirstColumn(firstColumn + insertCount);
                }
                address.setLastColumn(lastColumn + insertCount);
            }else if(lastColumn >= insertStart - 1){
                address.setLastColumn(lastColumn + insertCount);
            }
            excelBaseInfoDto.sheet.addMergedRegionUnsafe(address);
        }
        // 设置行高
        for (int i = 0; i < excelBaseInfoDto.headRowCount; i++) {
            excelBaseInfoDto.sheet.getRow(i).setHeight(excelBaseInfoDto.headAndEndRowHeightList.get(i));
        }
        excelBaseInfoDto.needSplit = true;
        // 初始化字段类型数据
        initFieldType(excelBaseInfoDto);
        return sheetIndex;
    }

    /**
     * 交换两个单元格格式及内容
     * @param cellA
     * @param cellB
     */
    private void exchangeCell(Cell cellA, Cell cellB) {
        String titleA = cellA.getStringCellValue();
        String titleB = cellB.getStringCellValue();
        CellStyle cellStyleA = cellA.getCellStyle();
        CellStyle cellStyleB = cellB.getCellStyle();
        cellA.setCellValue(titleB);
        cellA.setCellStyle(cellStyleB);
        cellB.setCellValue(titleA);
        cellB.setCellStyle(cellStyleA);
    }

    // 初始化字段类型数据
    private void initFieldType(ExcelBaseInfoDto excelBaseInfoDto) {

        if (excelBaseInfoDto.dataList == null || excelBaseInfoDto.dataList.isEmpty()) {
            return;
        }
        Class<?> clazz = excelBaseInfoDto.dataList.get(0).getClass();
        Field[] objfields = clazz.getDeclaredFields();
        clazz = clazz.getSuperclass();
        while (null != clazz) {
            Field[] superFields = clazz.getDeclaredFields();
            objfields = (Field[]) ArrayUtils.addAll(objfields, superFields);
            clazz = clazz.getSuperclass();
        }

        boolean isFind = false;
        for (FileDataInfo fileDataInfo : excelBaseInfoDto.bodyConfig) {
            isFind = false;
            for (Field f : objfields) {
                String fieldName = fileDataInfo.fieldName;
                if(excelBaseInfoDto.needSplit && !StringUtil.isNullOrEmpty(fieldName) && fieldName.contains(".")){
                    fieldName = fieldName.split("\\.")[0];
                }
                if (f.getName().equals(fieldName)) {
                    fileDataInfo.fieldType = f.getGenericType().toString();
                    isFind = true;
                    break;
                }
            }
            if (!isFind) {
                logger.warn(String.format("【%s】未找到对应数据源 ", fileDataInfo.fieldName));
                fileDataInfo.fieldType = this.emptyString;
                fileDataInfo.isErrorOrNoSuchMethod = true;
            }
        }
    }

    // 获取sheet的一行，没有则新增
    private Row getRow(int index, ExcelBaseInfoDto excelBaseInfoDto) {
        if (index <= excelBaseInfoDto.sheet.getLastRowNum()) {
            excelBaseInfoDto.sheet.shiftRows(index, excelBaseInfoDto.sheet.getLastRowNum(), 1, true, false);
        }
        Row row = excelBaseInfoDto.sheet.createRow(index);
        for (int i = 0; i < excelBaseInfoDto.colCount; i++) {
            if (row.getCell(i) == null) {
                row.createCell(i);
            }
        }
        if (excelBaseInfoDto.notSetLineHeight == 0) {
            row.setHeight(excelBaseInfoDto.rowHeight);
        }
        return row;
    }

    // 数据填充Excel行
    private void setRowFromObject(int index, ExcelBaseInfoDto excelBaseInfoDto)
            throws Exception {
        int rowNo = index + excelBaseInfoDto.headRowCount;
        // 获取一个行
        Row row = getRow(rowNo, excelBaseInfoDto);
        // 获取要输出的数据
        Object obj = excelBaseInfoDto.dataList.get(index);
        FileDataInfo mergeKeyCol = excelBaseInfoDto.mergeKeyCol;
        List<String> decimalPlaceControlList = excelBaseInfoDto.decimalPlaceControlList == null ? new ArrayList<>() : excelBaseInfoDto.decimalPlaceControlList;
        Map<Integer, CellStyle> subZeroStyleMap = excelBaseInfoDto.subZeroStyleMap == null ? new HashMap<>() : excelBaseInfoDto.subZeroStyleMap;
        Map<String, Integer> decimalPlaceNumMap = excelBaseInfoDto.getDecimalPlaceNumMap();
        for (FileDataInfo fileDataInfo : excelBaseInfoDto.bodyConfig) {
            Cell cell = row.getCell(fileDataInfo.cellIndex);
            if(null != fileDataInfo.cellStyle) {
                cell.setCellStyle(fileDataInfo.cellStyle);
            }
            if (!fileDataInfo.isErrorOrNoSuchMethod && !StringUtil.isNullOrEmpty(fileDataInfo.fieldName)) {
                String fieldName = fileDataInfo.fieldName;
                String secondName = "";
                String fields[] = fieldName.split("\\.");
                if(fields.length == 2){
                    fieldName = fields[0];
                    secondName = fields[1];
                }
                try {
                    Object value = null;
                    try {
                        value = getFieldValue(obj, fieldName);
                    } catch (NoSuchMethodException ex) {
                        fileDataInfo.isErrorOrNoSuchMethod = true;
                        logger.warn(String.format("【%s】未找到对应的【get】方法", fileDataInfo.fieldName), ex);
                    } catch (Exception ex) {
                        fileDataInfo.isErrorOrNoSuchMethod = true;
                        logger.warn(String.format("【%s】未找到对应的【get】方法:%s", fileDataInfo.fieldName, JSONObject.toJSONString(ex)));
                    }
                    if(fileDataInfo.isErrorOrNoSuchMethod){
                        continue;
                    }
                    if (fileDataInfo.mergeRow) {
                        Map<Object, List<CellRangeAddress>> mergeMap = excelBaseInfoDto.colMergeMap.get(fileDataInfo.cellIndex);
                        StringBuilder key = new StringBuilder("" + value);
                        // 兼容以前写法 判断mergeKeyCol
                        if (mergeKeyCol != null) {
                            Object mergeKey = getFieldValue(obj, mergeKeyCol.fieldName);
                            key.append("mergerKey").append(mergeKey);
                        }else{
                            for (String tmpFieldName : fileDataInfo.mergeFieldNameList) {
                                Object mergeKey = getFieldValue(obj, tmpFieldName);
                                key.append("mergerKey").append(mergeKey);
                            }
                        }
                        int colIndex = cell.getColumnIndex();
                        if (!mergeMap.containsKey(key.toString())) {
                            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowNo, rowNo, colIndex, colIndex);
                            List<CellRangeAddress> cellRangeAddressList = new ArrayList<>();
                            cellRangeAddressList.add(cellRangeAddress);
                            mergeMap.put(key.toString(), cellRangeAddressList);
                        } else {
                            List<CellRangeAddress> cellRangeAddressList = mergeMap.get(key.toString());
                            CellRangeAddress last = cellRangeAddressList.get(cellRangeAddressList.size() - 1);
                            if (last.getLastRow() + 1 != rowNo) {
                                // 相同值的数据不是上一行，说明是隔行了，需要是新的合并单元格
                                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowNo, rowNo, colIndex, colIndex);
                                cellRangeAddressList.add(cellRangeAddress);
                            } else {
                                last.setLastRow(rowNo);
                            }
                        }
                    }
                    if (value == null) {
                        continue;
                    }
                    switch (fileDataInfo.fieldType) {
                        case "class java.lang.String":
                            cell.setCellValue(value.toString());
                            break;
                        case "class java.util.Date":
                            String strValue = StringUtil.getFormatDateS((Date) value);
                            cell.setCellValue(strValue);
                            break;
                        case "class java.lang.Integer":
                        case "int":
                            cell.setCellValue((Integer) value);
                            break;
                        case "class java.lang.Boolean":
                        case "boolean":
                            cell.setCellValue((Boolean) value);
                            break;
                        case "class java.lang.Double":
                        case "double":
                            double doubleValue = (double) value;
                            if(decimalPlaceControlList.contains(fileDataInfo.fieldName) || decimalPlaceNumMap.containsKey(fileDataInfo.fieldName)){
                                int len = 6;
                                if(decimalPlaceControlList.contains(fileDataInfo.fieldName)) {
                                    len = getNumberDecimalDigits(doubleValue);
                                }
                                Integer lenParam = decimalPlaceNumMap.get(fileDataInfo.fieldName);
                                if (null != lenParam && lenParam < len) {
                                    len = lenParam;
                                }
                                CellStyle cellStyle = subZeroStyleMap.get(len);
                                if(null == cellStyle){
                                    cellStyle = excelBaseInfoDto.workbook.createCellStyle();
                                    cellStyle.cloneStyleFrom(cell.getCellStyle());
                                    DataFormat df = excelBaseInfoDto.workbook.createDataFormat();
                                    cellStyle.setDataFormat(df.getFormat(doubleDataFormatStr(len)));
                                    subZeroStyleMap.put(len, cellStyle);
                                }
                                cell.setCellStyle(cellStyle);
                            }
                            cell.setCellValue(doubleValue);
                            break;
                        case "class java.math.BigDecimal":
                            cell.setCellValue(((BigDecimal) value).doubleValue());
                            break;
                        case "class java.lang.Long":
                        case "long":
                            cell.setCellValue(((Long) value).toString());
                            break;
                        case "class java.lang.Byte":
                        case "byte":
                            cell.setCellValue((Byte) value);
                            break;
                        default:
                            if(!StringUtil.isNullOrEmpty(secondName) && fileDataInfo.fieldType.startsWith("java.util.Map")){
                                Object ob = ((Map<String, Double>) value).get(secondName);
                                doubleValue = Double.valueOf(String.valueOf(ob == null ? "" : ob));
                                if(decimalPlaceControlList.contains(fileDataInfo.fieldName) || decimalPlaceNumMap.containsKey(fileDataInfo.fieldName)){
                                    int len = getNumberDecimalDigits(doubleValue);
                                    Integer lenParam = decimalPlaceNumMap.get(fileDataInfo.fieldName);
                                    if(null != lenParam){
                                        len = lenParam;
                                    }
                                    CellStyle cellStyle = subZeroStyleMap.get(len);
                                    if(null == cellStyle){
                                        cellStyle = excelBaseInfoDto.workbook.createCellStyle();
                                        cellStyle.cloneStyleFrom(cell.getCellStyle());
                                        DataFormat df = excelBaseInfoDto.workbook.createDataFormat();
                                        cellStyle.setDataFormat(df.getFormat(doubleDataFormatStr(len)));
                                        subZeroStyleMap.put(len, cellStyle);
                                    }
                                    cell.setCellStyle(cellStyle);
                                }
                                cell.setCellValue(doubleValue);
                            }else {
                                logger.warn(String.format("【%s】不支持的字段类型:%s ", fileDataInfo.fieldName, fileDataInfo.fieldType));
                                fileDataInfo.isErrorOrNoSuchMethod = true;
                            }
                    }

                } catch (Exception ex) {
                    logger.warn(String.format("【%s】数据读取异常 ", fileDataInfo.fieldName), ex);
                }
            } else {
                cell.setCellValue(emptyString);
            }
        }
        if(excelBaseInfoDto.needMergedRegion){
            try {
                excelBaseInfoDto.sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, excelBaseInfoDto.firstColumn, excelBaseInfoDto.lastColumn));
            }catch (Exception ex){

            }
        }
        if (index == excelBaseInfoDto.dataList.size() - 1) {
            Map<Integer, Map<Object, List<CellRangeAddress>>> colMergeMap = excelBaseInfoDto.colMergeMap;
            for (Map<Object, List<CellRangeAddress>> mergeMap : colMergeMap.values()) {
                for (List<CellRangeAddress> cellList : mergeMap.values()) {
                    for (CellRangeAddress cell : cellList) {
                        int firstRow = cell.getFirstRow();
                        int lastRow = cell.getLastRow();
                        int firstColumn = cell.getFirstColumn();
                        int lastColumn = cell.getLastColumn();
                        if (firstRow != lastRow || firstColumn != lastColumn) {
                            excelBaseInfoDto.sheet.addMergedRegion(cell);
                        }
                    }
                }
            }
        }
    }

    private Object getFieldValue(Object obj, String field) throws NoSuchMethodException {
        String pre = field.substring(0, 1);
        String methodName = "get" + pre.toUpperCase() + field.substring(1);
        Class<?> clazz = obj.getClass();
        Object value = null;
        while (clazz != Object.class) {
            try {
                Method method = clazz.getDeclaredMethod(methodName);
                value = method.invoke(obj);
                break;
            } catch (NoSuchMethodException ex) {
                clazz = clazz.getSuperclass();
                if(clazz == Object.class) {
                    throw ex;
                }
                try{
                    methodName = "get" + field;
                    Method method = clazz.getDeclaredMethod(methodName);
                    value = method.invoke(obj);
                    break;
                }catch (NoSuchMethodException ex2){
                    clazz = clazz.getSuperclass();
                    if(clazz == Object.class) {
                        throw ex2;
                    }
                } catch (Exception ex2) {
                    break;
                }
            } catch (Exception ex) {
                break;
            }
        }
        return value;
    }

    private static int getNumberDecimalDigits(double number) {
        int refValue = 0;
        if (number % 1 != 0) {
            int i = 0;
            while (i < 6) {
                i++;
                if (number * Math.pow(10, i) % 1 == 0) {
                    break;
                }
            }
            refValue = i;
        }
        return refValue;
    }

    private static String doubleDataFormatStr(int len){
        String refStr = "#,##0";
        if(len != 0){
            refStr = String.format("#,##0.%s", "0000000000".substring(0,len));
        }
        return refStr;
    }

    // 数据字段配置信息
    class FileDataInfo {
        FileDataInfo(int cellIndex, String fieldName) {
            this.cellIndex = cellIndex;
            this.fieldName = fieldName;
        }

        FileDataInfo(int rowIndex, int cellIndex, String fieldName) {
            this.rowIndex = rowIndex;
            this.cellIndex = cellIndex;
            this.fieldName = fieldName;
        }

        /** 模板中对应的行位置 */
        int rowIndex;
        /** 模板中对应的列位置 */
        int cellIndex;
        /** 模板中设置的字段名称-对应DTO属性 */
        String fieldName;
        /** DTO对应属性类型 */
        String fieldType;
        /** 在DTO找get属性是否异常 */
        boolean isErrorOrNoSuchMethod;
        /** 数据单元格格式 */
        CellStyle cellStyle;
        /** 此列是否合并单元格 */
        boolean mergeRow = false;
        /** 此列是否合并单元格 */
        List<String> mergeFieldNameList = new ArrayList<>();
    }

    class ExcelBaseInfoDto implements Cloneable {
        ExcelBaseInfoDto() {
            dataList = new ArrayList<>();
            headAndEbdConfig = new ArrayList<>();
            bodyConfig = new ArrayList<>();
            headAndEndRowHeightList = new ArrayList<>();
        }

        // Excel数据
        protected List<?> dataList;
        // Excel配置信息
        protected List<FileDataInfo> headAndEbdConfig;
        protected List<FileDataInfo> bodyConfig;
        // 模板文件名称（带路径）
        protected String fieldFullName;
        // 模板文件名称（带路径）
        protected InputStream inputStream;
        // Excel模板Book
        protected Workbook workbook;
        // Excel模板Sheet
        protected Sheet sheet;
        // 头信息行数
        protected int headRowCount;
        // 数据总列数
        protected int colCount;
        // 表头变动数据数
        protected int headConfigCount;
        // 尾信息行数
        protected int endRowCount;
        // 表尾变动数据数
        protected int endConfigCount;
        // 配置信息行数
        protected final int configureRowCount = 3;
        // 表头动态列数量
        protected int headUnFixColCount;
        // 表头动态列行号
        protected int unFixColRowNo;

        /** 行高 */
        protected short rowHeight;

        /** 表头表尾行高 */
        protected List<Short> headAndEndRowHeightList;

        /** 需要合并 */
        protected boolean needMergedRegion;

        /** 需要分割下级字段名 */
        protected boolean needSplit;

        /** 合并开始单元格 */
        protected int firstColumn;

        /** 合并结束单元格 */
        protected int lastColumn;

        /**
         * 是否复制设置行行高
         */
        int notSetLineHeight = 0;

        /**
         * 小数位控制字段
         */
        List<String> decimalPlaceControlList;

        /**
         * 小数位控制字典
         * 不设置时按去尾零控制
         */
        private Map<String, Integer> decimalPlaceNumMap;

        public Map<String, Integer> getDecimalPlaceNumMap() {
            if(this.decimalPlaceNumMap == null){
                this.decimalPlaceNumMap = new HashMap<>();
            }
            return decimalPlaceNumMap;
        }

        public void setDecimalPlaceNumMap(Map<String, Integer> decimalPlaceNumMap) {
            this.decimalPlaceNumMap = decimalPlaceNumMap;
        }

        /** 页签索引，用于子类模板 */
        int sheetIndex;
        /** 子类模板集合 */
        protected List<ExcelBaseInfoDto> sonList;

        protected Map<Integer, CellStyle> subZeroStyleMap = new HashMap<>();

        /** 需要合并的单元格信息，以列序号、标识列(如果有)+合并列单元格内容为 key */
        protected Map<Integer, Map<Object, List<CellRangeAddress>>> colMergeMap = new HashMap<>();

        /** 合并单元格时作为分组的标识列*/
        protected FileDataInfo mergeKeyCol;

        public ExcelBaseInfoDto clone(int sheetIndex) {
            ExcelBaseInfoDto dto = new ExcelBaseInfoDto();
            dto.headAndEbdConfig = this.headAndEbdConfig;
            dto.headAndEndRowHeightList = this.headAndEndRowHeightList;
            dto.bodyConfig = this.bodyConfig;
            dto.needMergedRegion = this.needMergedRegion;
            dto.colMergeMap = this.colMergeMap;
            dto.mergeKeyCol = this.mergeKeyCol;
            dto.sheetIndex = sheetIndex;
            return dto;
        }

        /** 根据页签索引，取对应的头、体模板 */
        public void setConfigBySon(int sheetIndex) {
            ExcelBaseInfoDto sonDto = this.sonList.stream().filter(son -> {return son.sheetIndex == sheetIndex;}).collect(Collectors.toList()).get(0);
            if (sonDto != null) {
                this.headAndEbdConfig = sonDto.headAndEbdConfig;
                this.headAndEndRowHeightList = sonDto.headAndEndRowHeightList;
                this.bodyConfig = sonDto.bodyConfig;
                this.headRowCount = sonDto.headAndEndRowHeightList.size();
                this.colCount = sonDto.bodyConfig.size();
                this.colMergeMap = sonDto.colMergeMap;
            }
        }

        /** 清理配置信息 */
        public void clearConfig() {
            this.headAndEbdConfig = new ArrayList<>();
            this.headAndEndRowHeightList = new ArrayList<>();
            this.bodyConfig = new ArrayList<>();
            this.colMergeMap = new HashMap<>();
        }

        public void clearColMergeMap() {
            for (Integer key : colMergeMap.keySet()) {
                colMergeMap.put(key, new HashMap<>());
            }
        }

    }
}
