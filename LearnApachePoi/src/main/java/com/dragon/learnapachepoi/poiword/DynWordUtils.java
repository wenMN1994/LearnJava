package com.dragon.learnapachepoi.poiword;

//import junit.framework.Assert;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.*;

/**
 * word文档模板文件生成
 * 1:绘制表格(指定位置插入,默认表格末尾新增)2:绘制段落文字内容3:标签处替换图片等
 *
 * @author DragonWen
 */
public class DynWordUtils {
	private static Logger logger = LoggerFactory.getLogger(DynWordUtils.class);
	
	/**
     * 被list替换的段落 被替换的都是oldParagraph
     */
    private XWPFParagraph oldParagraph;

    /**
     * 参数
     */
    private Map<String, Object> paramMap;

    /**
     * 当前元素的位置
     */
    int n = 0;

    /**
     * 判断当前是否是遍历的表格
     */
    boolean isTable = false;

    /**
     * 模板对象
     */
    XWPFDocument templateDoc;
    
    /**
     * 默认字体的大小
     */
    final int DEFAULT_FONT_SIZE = 10;

    /**
     * 重复模式的占位符所在的行索引
     */
    private int currentRowIndex;
    
    
    private int currentTableIndex;// 当前表格
    
    private int currentTableTotalIndex;// 当前表格总行数（注意动态表格是在指定位置新增单元格，行数对应新增）

    /**
     * 入口
     * @param paramMap     模板中使用的参数
     * @param templatePaht 模板全路径
     * @param outPath      生成的文件存放的本地全路径
     */
    public static void process(Map<String, Object> paramMap, String templatePaht, String outPath) {
        // 入口
    	DynWordUtils dynWordUtils = new DynWordUtils();
        dynWordUtils.setParamMap(paramMap);
        dynWordUtils.createWord(templatePaht, outPath);
    }

    /**
     * 生成动态的word
     * @param templatePath
     * @param outPath
     */
    public void createWord(String templatePath, String outPath) {
        File inFile = new File(templatePath);
        if(!inFile.exists()){
        	throw new RuntimeException("未获取到模板文件");
        }
        
        InputStream inputStream = null;
        OutputStream outputStream = null;
        try {
	    	inputStream = new FileInputStream(new File(templatePath));
	    	outputStream = new FileOutputStream(outPath);
	    	// templateDoc = new XWPFDocument(OPCPackage.open(inFile));
	    	templateDoc = new XWPFDocument(inputStream);
            parseTemplateWord();
            templateDoc.write(outputStream);
            
        } catch (Exception e) {
			logger.error(e.getMessage(),e);
			StackTraceElement[] stackTrace = e.getStackTrace();
            String className = stackTrace[0].getClassName();
            String methodName = stackTrace[0].getMethodName();
            int lineNumber = stackTrace[0].getLineNumber();
            logger.error("错误：第:{}行, 类名:{}, 方法名:{}", lineNumber, className, methodName);
            logger.error("word文档文件生成异常createWord:"+e.getMessage(), e);;	
            throw new RuntimeException(e.getCause().getMessage());
		}finally{
			try {
				if (inputStream != null) {
					inputStream.close();
				}
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			}
			try {
				if (outputStream != null) {
					outputStream.close();
				}
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			}
		}
        
       /* try (FileOutputStream outStream = new FileOutputStream(outPath)) {
            templateDoc = new XWPFDocument(OPCPackage.open(inFile));
            parseTemplateWord();
            templateDoc.write(outStream);
        } catch (Exception e) {
            StackTraceElement[] stackTrace = e.getStackTrace();
            String className = stackTrace[0].getClassName();
            String methodName = stackTrace[0].getMethodName();
            int lineNumber = stackTrace[0].getLineNumber();
            logger.error("错误：第:{}行, 类名:{}, 方法名:{}", lineNumber, className, methodName);
            logger.error("word文档文件生成异常createWord:"+e.getMessage(), e);;	
            throw new RuntimeException(e.getCause().getMessage());
        }*/
    }

    /**
     * 解析word模板
     */
    public void parseTemplateWord() throws Exception {
        List<IBodyElement> elements = templateDoc.getBodyElements();
        for (; n < elements.size(); n++) {
            IBodyElement element = elements.get(n);
            // 解析word文档内容：普通段落
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                oldParagraph = paragraph;
                if (paragraph.getParagraphText().isEmpty()) {
                    continue;
                }
                delParagraph(paragraph);
            } else if (element instanceof XWPFTable) {
                // 解析word文档内容：表格段落
                isTable = true;
                XWPFTable table = (XWPFTable) element;// 当前表格对象
                settleTable(table, paramMap);
                isTable = false;
            }
        }
    }

    /**
     * 处理段落
     */
    private void delParagraph(XWPFParagraph paragraph) throws Exception {
        List<XWPFRun> runs = oldParagraph.getRuns();
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text == null) {
                continue;
            }
            sb.append(text);
            run.setText("", 0);
        }
        Placeholder(paragraph, runs, sb);
    }


    /**
     * 匹配传入信息集合与模板
     * @param placeholder 模板需要替换的区域()
     * @param paramMap    传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public void changeValue(XWPFRun currRun, String placeholder, Map<String, Object> paramMap) throws Exception {
    	
        String placeholderValue = placeholder;
        if (paramMap == null || paramMap.isEmpty()) {
            return;
        }
        Set<Map.Entry<String, Object>> textSets = paramMap.entrySet();
        for (Map.Entry<String, Object> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String mapKey = textSet.getKey();
            String docKey = PoiWordUtils.getDocKey(mapKey);

            if (placeholderValue.indexOf(docKey) != -1) {
                Object obj = textSet.getValue();
                // 需要添加一个list
                if (obj instanceof List) {
                    placeholderValue = delDynList(placeholder, (List) obj);
                } else {
                    placeholderValue = placeholderValue.replaceAll(
                            PoiWordUtils.getPlaceholderReg(mapKey)
                            , String.valueOf(obj));
                }
            }
        }

        currRun.setText(placeholderValue, 0);
    }

    /**
     * 处理的动态的段落（参数为list）
     *
     * @param placeholder 段落占位符
     * @param obj
     * @return
     */
    private String delDynList(String placeholder, List obj) {
        String placeholderValue = placeholder;
        List dataList = obj;
        Collections.reverse(dataList);
        for (int i = 0, size = dataList.size(); i < size; i++) {
            Object text = dataList.get(i);
            // 占位符的那行, 不用重新创建新的行
            if (i == 0) {
                placeholderValue = String.valueOf(text);
            } else {
                XWPFParagraph paragraph = createParagraph(String.valueOf(text));
                if (paragraph != null) {
                    oldParagraph = paragraph;
                }
                // 增加段落后doc文档会的element的size会随着增加（在当前行的上面添加），回退并解析新增的行（因为可能新增的带有占位符，这里为了支持图片和表格）
                if (!isTable) {
                    n--;
                }
            }
        }
        return placeholderValue;
    }

    /**
     * 创建段落
     * @param texts
     */
    public XWPFParagraph createParagraph(String... texts) {
        // 使用游标创建一个新行
        XmlCursor cursor = oldParagraph.getCTP().newCursor();
        XWPFParagraph newPar = templateDoc.insertNewParagraph(cursor);
        // 设置段落样式
        newPar.getCTP().setPPr(oldParagraph.getCTP().getPPr());
        copyParagraph(oldParagraph, newPar, texts);
        return newPar;
    }

    /**
     * 处理当前表格（遍历）
     * @param table    表格
     * @param paramMap 需要替换的信息集合
     */
    public void settleTable(XWPFTable table, Map<String, Object> paramMap) throws Exception {
        // 处理当前表格
    	currentTableIndex++;
    	String messagePre = "当前解析表格"+currentTableIndex; 
    	logger.info(messagePre);
    	List<XWPFTableRow> rows = table.getRows();
    	logger.info("当前表格行数共计="+rows.size());
    	currentTableTotalIndex = rows.size();
    	int totalRows = rows.size();
    	
        for (int i = 0; i < totalRows; i++) {// 遍历当前表格段落的每一行
            XWPFTableRow row = rows.get(i);
            totalRows = rows.size();
            logger.info("解析后当前表格行数共计="+totalRows);
            currentRowIndex = i;
            logger.info(messagePre+"遍历当前行数赋值="+currentRowIndex);
            // 如果是动态添加行 直接处理后，如果是插入动态位置，其后面的row需要继续解析
            if (delAndJudgeRow(table, paramMap, row)) {
                logger.info("动态行处理完成");
            	//return;
            }
        }
        logger.info("currentTableTotalIndex="+currentTableTotalIndex);
    }

    /**
     * 判断并且是否是动态行，并且处理表格占位符
     * @param table 表格对象
     * @param paramMap 参数map
     * @param row 当前行
     * @return
     * @throws Exception
     */
    private boolean delAndJudgeRow(XWPFTable table, Map<String, Object> paramMap, XWPFTableRow row) throws Exception {
        // 表格动态行处理
        logger.info("当前行数="+currentRowIndex);
    	if (PoiWordUtils.isAddRow(row)) {
            List<XWPFTableRow> xwpfTableRows = addAndGetRows(table, row, paramMap);
            // 回溯添加的行，这里是试图处理动态添加的图片
            for (XWPFTableRow tbRow : xwpfTableRows) {
                delAndJudgeRow(table, paramMap, tbRow);
            }
            return true;
        }

        // 如果是重复添加的行
        if (PoiWordUtils.isAddRowRepeat(row)) {
            List<XWPFTableRow> xwpfTableRows = addAndGetRepeatRows(table, row, paramMap);
            // 回溯添加的行，这里是试图处理动态添加的图片
            for (XWPFTableRow tbRow : xwpfTableRows) {
                delAndJudgeRow(table, paramMap, tbRow);
            }
            return true;
        }
        
        // 如果是合并行列情况
        if (PoiWordUtils.isMergeRow(row)) {
        	List<XWPFTableRow> xwpfTableRows = addAndGetMergeRows(table, row, paramMap);
        	// 回溯添加的行，这里是试图处理动态添加的图片
        	for (XWPFTableRow tbRow : xwpfTableRows) {
        		delAndJudgeRow(table, paramMap, tbRow);
        	}
        	return true;
        }
        
        // 当前行非动态行标签
        List<XWPFTableCell> cells = row.getTableCells();
        for (XWPFTableCell cell : cells) {
            // 判断单元格是否需要替换
            if (PoiWordUtils.checkText(cell.getText())) {
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    List<XWPFRun> runs = paragraph.getRuns();
                    StringBuilder sb = new StringBuilder();
                    for (XWPFRun run : runs) {
                        sb.append(run.toString());
                        run.setText("", 0);
                    }
                    Placeholder(paragraph, runs, sb);
                }
            }
        }
        return false;
    }

    /**
     * 处理占位符
     * @param runs 当前段的runs
     * @param sb 当前段的内容
     * @throws Exception
     */
    private void Placeholder(XWPFParagraph currentPar, List<XWPFRun> runs, StringBuilder sb) throws Exception {
        if (runs.size() > 0) {
            String text = sb.toString();
            XWPFRun currRun = runs.get(0);
            if (PoiWordUtils.isPicture(text)) {
                // 该段落是图片占位符
                ImageEntity imageEntity = (ImageEntity) PoiWordUtils.getValueByPlaceholder(paramMap, text);
                int indentationFirstLine = currentPar.getIndentationFirstLine();
                // 清除段落的格式，否则图片的缩进有问题
                currentPar.getCTP().setPPr(null);
                //设置缩进
                currentPar.setIndentationFirstLine(indentationFirstLine);
                addPicture(currRun, imageEntity);
            } else {
                changeValue(currRun, text, paramMap);
            }
        }
    }

    /**
     * 添加图片
     * @param currRun 当前run
     * @param imageEntity 图片对象
     * @throws InvalidFormatException
     * @throws FileNotFoundException
     */
    private void addPicture(XWPFRun currRun, ImageEntity imageEntity) throws InvalidFormatException, FileNotFoundException {
        if(imageEntity != null){
        	Integer typeId = imageEntity.getTypeId().getTypeId();
            String picId = currRun.getDocument().addPictureData(new FileInputStream(imageEntity.getUrl()), typeId);
            ImageUtils.createPicture(currRun, picId, templateDoc.getNextPicNameNumber(typeId),
                    imageEntity.getWidth(), imageEntity.getHeight());
        }
    }

    /**
     * 添加行  标签行不是新创建的
     *
     * @param table
     * @param flagRow  flagRow 表有标签的行
     * @param paramMap 参数
     */
    private List<XWPFTableRow> addAndGetRows(XWPFTable table, XWPFTableRow flagRow, Map<String, Object> paramMap) throws Exception {
        List<XWPFTableCell> flagRowCells = flagRow.getTableCells();
        XWPFTableCell flagCell = flagRowCells.get(0);
        
        String text = flagCell.getText();// ${tbAddRow:tb1}
        @SuppressWarnings("unchecked")
		List<List<String>> dataList = (List<List<String>>) PoiWordUtils.getValueByPlaceholder(paramMap, text);
        if(dataList == null){
        	return new ArrayList<XWPFTableRow>();
        }
        
        // 新添加的行
        List<XWPFTableRow> newRows = new ArrayList<>(dataList.size());
        if (dataList == null || dataList.size() <= 0) {
            return newRows;
        }

        XWPFTableRow currentRow = flagRow;
        int cellSize = flagRow.getTableCells().size();
        for (int i = 0, size = dataList.size(); i < size; i++) {
            if (i != 0) {
                currentRow = table.createRow();// 在表格最后加行，不适用于复杂表格

                // 复制样式
                if (flagRow.getCtRow() != null) {
                    currentRow.getCtRow().setTrPr(flagRow.getCtRow().getTrPr());
                }
            }
            addRow(flagCell, currentRow, cellSize, dataList.get(i));
            newRows.add(currentRow);
        }
        return newRows;
    }

    /**
     * 添加重复多行 动态行  每一行都是新创建的
     * @param table
     * @param flagRow
     * @param paramMap
     * @return
     * @throws Exception
     */
    private List<XWPFTableRow> addAndGetRepeatRows(XWPFTable table, XWPFTableRow flagRow, Map<String, Object> paramMap) throws Exception {
        List<XWPFTableCell> flagRowCells = flagRow.getTableCells();
        XWPFTableCell flagCell = flagRowCells.get(0);
        String text = flagCell.getText();
        @SuppressWarnings("unchecked")
		List<List<String>> dataList = (List<List<String>>) PoiWordUtils.getValueByPlaceholder(paramMap, text);
        String tbRepeatMatrix = PoiWordUtils.getTbRepeatMatrix(text);
//        Assert.assertNotNull("模板矩阵不能为空", tbRepeatMatrix);
        
        if(dataList == null){
        	return new ArrayList<XWPFTableRow>();
        }
        
        // 新添加的行
        List<XWPFTableRow> newRows = new ArrayList<>(dataList.size());
        if (dataList == null || dataList.size() <= 0) {
            return newRows;
        }

        String[] split = tbRepeatMatrix.split(PoiWordUtils.tbRepeatMatrixSeparator);
        int startRow = Integer.parseInt(split[0]);
        int endRow = Integer.parseInt(split[1]);
        int startCell = Integer.parseInt(split[2]);
        int endCell = Integer.parseInt(split[3]);

        XWPFTableRow currentRow;
        for (int i = 0, size = dataList.size(); i < size; i++) {
            int flagRowIndex = i % (endRow - startRow + 1);
            XWPFTableRow repeatFlagRow = table.getRow(flagRowIndex);
            // 清除占位符那行
            if (i == 0) {
                table.removeRow(currentRowIndex);
            }
            currentRow = table.createRow();// 在表格最后加行，不适用于复杂表格
            // 复制样式
            if (repeatFlagRow.getCtRow() != null) {
                currentRow.getCtRow().setTrPr(repeatFlagRow.getCtRow().getTrPr());
            }
            addRowRepeat(startCell, endCell, currentRow, repeatFlagRow, dataList.get(i));
            newRows.add(currentRow);
        }
        return newRows;
    }
    
    /**
     * 添加合并行列等
     * @param table
     * @param flagRow
     * @param paramMap
     * @return
     * @throws Exception
     */
    private List<XWPFTableRow> addAndGetMergeRows(XWPFTable table, XWPFTableRow flagRow, Map<String, Object> paramMap) throws Exception {
    	// 添加动态行并且合并等
    	List<XWPFTableCell> flagRowCells = flagRow.getTableCells();// 获取待添加的数据总数
    	XWPFTableCell flagCell = flagRowCells.get(0);
    	
    	String text = flagCell.getText();// ${tbAddRowMerge:tb3}
    	// List<List<String>> dataList = (List<List<String>>) PoiWordUtils.getValueByPlaceholder(paramMap, text);
    	@SuppressWarnings("unchecked")
		List<Map<String,List<List<String>>>> dataAddList = (List<Map<String, List<List<String>>>>) PoiWordUtils.getValueByPlaceholder(paramMap, text);
    	
    	int totalRow = 0;
    	if(CollectionUtils.isNotEmpty(dataAddList)){
    		for (Map<String, List<List<String>>> map : dataAddList) {
    			Set<String> keySet = map.keySet();
    			for (String key : keySet) {
    				if(StringUtils.isBlank(key)){
    					continue;
    				}
    				List<List<String>> dataList = map.get(key);
    				if(dataList == null){
    					continue;
    				}
    				totalRow = totalRow + dataList.size();
    			}
    		}
    	}
    	// 新添加的行
    	List<XWPFTableRow> newRows = new ArrayList<>(totalRow);
    	if(totalRow == 0){
    		return newRows;
    	}
    	XWPFTableRow currentRow = flagRow;
    	int cellSize = flagRow.getTableCells().size();// 待添加数据行数
    	int totalIndex = 0;// 总添加数据计数
    	if(CollectionUtils.isNotEmpty(dataAddList)){
    		for (Map<String, List<List<String>>> map : dataAddList) {
    			Set<String> keySet = map.keySet();
    			for (String key : keySet) {
    				if(StringUtils.isBlank(key)){
    					continue;
    				}
    		    	List<List<String>> dataList = map.get(key);
    		    	
    				boolean rowFlag = key.contains("tbAddRowMergeRow");// 合并行
    				boolean colFlag = key.contains("tbAddRowMergeCol");// 合并列
    				
    				if (dataList == null || dataList.size() <= 0) {
    		    		return newRows;
    		    	}
    		    	for (int i = 0, size = dataList.size(); i < size; i++) {
    		    		if (totalIndex != 0) {
    		                // currentRow = table.createRow();// 在表格最后加行，不适用于复杂表格
    		                currentRow = table.insertNewTableRow(currentRowIndex+totalIndex);// 指定行添加
    		                // 复制样式
    		                if (flagRow.getCtRow() != null) {
    		                    currentRow.getCtRow().setTrPr(flagRow.getCtRow().getTrPr());
    		                }
    		                currentTableTotalIndex++;// 总行数新增
    		            }
    		    		
    		    		// flagCell:模板列(标记占位符的那个cell),row:新增的行,cellSize:每行的列数量（用来补列补足的情况）；rowDataList 每行的数据
    		    		insertRow(flagCell, currentRow, cellSize, dataList.get(i));
    		            
    		            newRows.add(currentRow);// 注意这儿添加行，会直接从最后一列开始添加
    		            totalIndex++;
    		    	}
    		    	
    		    	// 合并
    		    	String lastStr = "";
    				String[] splitArr = key.split(":");
    				if(splitArr != null && splitArr.length > 1){
    					lastStr = splitArr[1];
    				}
    				String[] split = lastStr.split(",");
    		    	int startRowCol = Integer.parseInt(split[0]);// 合并行：合并开始列；合并列：合并开始行
    		    	int start = Integer.parseInt(split[1]);// 开始行/列
    		    	int last = Integer.parseInt(split[2]);// 结束行/列
    		    	// int endCell = Integer.parseInt(split[3]);
    				if(rowFlag){// 合并行
    					mergeCellVertically(table, startRowCol, start, last);
    				}else if(colFlag){// 合并列
    					mergeCellsHorizontal(table, startRowCol, start, last);
    				}
    			}
			}
    	}
    	return newRows;
    }

    /**
     * 根据模板cell添加新行
     *
     * @param flagCell    模板列(标记占位符的那个cell)
     * @param row         新增的行
     * @param cellSize    每行的列数量（用来补列补足的情况）
     * @param rowDataList 每行的数据
     */
    private void addRow(XWPFTableCell flagCell, XWPFTableRow row, int cellSize, List<String> rowDataList) {
        for (int i = 0; i < cellSize; i++) {
            XWPFTableCell cell = row.getCell(i);
            cell = cell == null ? row.createCell() : row.getCell(i);
            if (i < rowDataList.size()) {
                PoiWordUtils.copyCellAndSetValue(flagCell, cell, rowDataList.get(i));
            } else {
                // 数据不满整行时，添加空列
                PoiWordUtils.copyCellAndSetValue(flagCell, cell, "");
            }
        }
    }
    
    /**
   	 * insertRow 在word表格中指定位置插入一行，并将某一行的样式复制到新增行
   	 * @param copyrowIndex 需要复制的行位置
   	 * @param newrowIndex 需要新增一行的位置
   	 * */
   	public static void insertRowPosition(XWPFTable table, int copyrowIndex, int newrowIndex) {
   		// 在表格中指定的位置新增一行
   		XWPFTableRow targetRow = table.insertNewTableRow(newrowIndex);
   		// 获取需要复制行对象
   		XWPFTableRow copyRow = table.getRow(copyrowIndex);
   		//复制行对象
   		targetRow.getCtRow().setTrPr(copyRow.getCtRow().getTrPr());
   		//或许需要复制的行的列
   		List<XWPFTableCell> copyCells = copyRow.getTableCells();
   		//复制列对象
   		XWPFTableCell targetCell = null;
   		for (int i = 0; i < copyCells.size(); i++) {
   			XWPFTableCell copyCell = copyCells.get(i);
   			targetCell = targetRow.addNewTableCell();
   			targetCell.getCTTc().setTcPr(copyCell.getCTTc().getTcPr());
   			if (copyCell.getParagraphs() != null && copyCell.getParagraphs().size() > 0) {
   				targetCell.getParagraphs().get(0).getCTP().setPPr(copyCell.getParagraphs().get(0).getCTP().getPPr());
   				if (copyCell.getParagraphs().get(0).getRuns() != null
   						&& copyCell.getParagraphs().get(0).getRuns().size() > 0) {
   					XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
   					cellR.setBold(copyCell.getParagraphs().get(0).getRuns().get(0).isBold());
   				} 
   			} 
   		}
   	}
    
   	/**
   	 * 新增行
   	 * @param flagCell
   	 * @param row
   	 * @param cellSize
   	 * @param rowDataList
   	 */
    private void insertRow(XWPFTableCell flagCell, XWPFTableRow row, int cellSize, List<String> rowDataList) {
    	for (int i = 0; i < cellSize; i++) {
    		XWPFTableCell cell = row.getCell(i);
    		cell = cell == null ? row.createCell() : row.getCell(i);
    		if (i < rowDataList.size()) {
    			PoiWordUtils.copyCellAndSetValue(flagCell, cell, rowDataList.get(i));
    		} else {
    			// 数据不满整行时，添加空列
    			PoiWordUtils.copyCellAndSetValue(flagCell, cell, "");
    		}
    	}
    }

    /**
     * 根据模板cell  添加重复行
     * @param startCell 模板列的开始位置
     * @param endCell 模板列的结束位置
     * @param currentRow 创建的新行
     * @param repeatFlagRow 模板列所在的行
     * @param rowDataList 每行的数据
     */
    private void addRowRepeat(int startCell, int endCell, XWPFTableRow currentRow, XWPFTableRow repeatFlagRow, List<String> rowDataList) {
        int cellSize = repeatFlagRow.getTableCells().size();
        for (int i = 0; i < cellSize; i++) {
            XWPFTableCell cell = currentRow.getCell(i);
            cell = cell == null ? currentRow.createCell() : currentRow.getCell(i);
            int flagCellIndex = i % (endCell - startCell + 1);
            XWPFTableCell repeatFlagCell = repeatFlagRow.getCell(flagCellIndex);
            if (i < rowDataList.size()) {
                PoiWordUtils.copyCellAndSetValue(repeatFlagCell, cell, rowDataList.get(i));
            } else {
                // 数据不满整行时，添加空列
                PoiWordUtils.copyCellAndSetValue(repeatFlagCell, cell, "");
            }
        }
    }

    /**
     * 复制段落
     *
     * @param sourcePar 原段落
     * @param targetPar
     * @param texts
     */
    private void copyParagraph(XWPFParagraph sourcePar, XWPFParagraph targetPar, String... texts) {

        targetPar.setAlignment(sourcePar.getAlignment());
        targetPar.setVerticalAlignment(sourcePar.getVerticalAlignment());

        // 设置布局
        targetPar.setAlignment(sourcePar.getAlignment());
        targetPar.setVerticalAlignment(sourcePar.getVerticalAlignment());

        if (texts != null && texts.length > 0) {
            String[] arr = texts;
            XWPFRun xwpfRun = sourcePar.getRuns().size() > 0 ? sourcePar.getRuns().get(0) : null;

            for (int i = 0, len = texts.length; i < len; i++) {
                String text = arr[i];
                XWPFRun run = targetPar.createRun();

                run.setText(text);

                run.setFontFamily(xwpfRun.getFontFamily());
                int fontSize = xwpfRun.getFontSize();
                run.setFontSize((fontSize == -1) ? DEFAULT_FONT_SIZE : fontSize);
                run.setBold(xwpfRun.isBold());
                run.setItalic(xwpfRun.isItalic());
            }
        }
    }

    public void setParamMap(Map<String, Object> paramMap) {
        this.paramMap = paramMap;
    }
    
    /**
     * word单元格行合并
     * @param table
     * @param col 需要合并的列
     * @param fromRow 开始行
     * @param toRow 结束行
     */
    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for(int rowIndex = fromRow; rowIndex <= toRow; rowIndex++){
            try {
				CTVMerge vmerge = CTVMerge.Factory.newInstance();
				if(rowIndex == fromRow){
				    vmerge.setVal(STMerge.RESTART);
				} else {
				    vmerge.setVal(STMerge.CONTINUE);
				}
				XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
				if(cell != null){
					CTTcPr tcPr = cell.getCTTc().getTcPr();
				    if (tcPr != null) {
				        tcPr.setVMerge(vmerge);
				    } else {
				        tcPr = CTTcPr.Factory.newInstance();
				        tcPr.setVMerge(vmerge);
				        cell.getCTTc().setTcPr(tcPr);
				    }
				}
			} catch (Exception e) {
				logger.error("mergeCellVertically异常:"+e.getMessage(),e);
			}
        }
    }
    
	/**
	 * word单元格列合并
	 * @param table 表格
	 * @param row 合并列所在行
	 * @param startCell 开始列
	 * @param endCell 结束列
	 */
	public static void mergeCellsHorizontal(XWPFTable table, int row, int startCell, int endCell) {
		try {
			for (int i = startCell; i <= endCell; i++) {
				XWPFTableCell cell = table.getRow(row).getCell(i);
				if (i == startCell) {
					// The first merged cell is set with RESTART merge value  
					cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
				} else {
					// Cells which join (merge) the first one, are set with CONTINUE  
					cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
				}
			}
		} catch (Exception e) {
			logger.error("mergeCellsHorizontal异常:"+e.getMessage(),e);
		}
	}
    
}
