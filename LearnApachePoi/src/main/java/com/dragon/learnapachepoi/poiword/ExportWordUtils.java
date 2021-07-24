package com.dragon.learnapachepoi.poiword;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 导出Word文档
 * https://www.cnblogs.com/xueershewang/p/11325756.html
 * @author DragonWen
 */
public class ExportWordUtils extends XWPFDocument {
	private static Logger logger = LoggerFactory.getLogger(ExportWordUtils.class);
	
	public static final String TYPE_PARAGRAPH = "type_paragraph";// 生成段落
	public static final String TYPE_IMG = "type_img";// 生成图片
	
	/**
	 * 只创建段落文章的doc
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public static XWPFDocument createParagraphDoc(List<Map<String, Object>> list) throws Exception{
		XWPFDocument document = new XWPFDocument();
		// 设置边距
		//CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		//CTPageMar pageMar = sectPr.addNewPgMar();
		//pageMar.setLeft(BigInteger.valueOf(720L));
		//pageMar.setTop(BigInteger.valueOf(1440L));
		//pageMar.setRight(BigInteger.valueOf(720L));
		//pageMar.setBottom(BigInteger.valueOf(1440L));
		
		for (Map<String, Object> map : list) {
			String paragraphContent = map.get(TYPE_PARAGRAPH) == null ? "无" : map.get(TYPE_PARAGRAPH)+"";// 段落标志
			
			// ==================== 设置段落 ======================
			XWPFParagraph firstParagraph = document.createParagraph();
			firstParagraph.setAlignment(ParagraphAlignment.LEFT);// 设置段落内容居左
			// firstParagraph.setIndentationRight(500);// 末尾缩进300
			
			paragraphContent = paragraphContent.replaceAll("\n", "\r");// 将写入字符的换行符换成word识别换行符
			
			XWPFRun run = firstParagraph.createRun();
			run.setText(paragraphContent);// 写入段落内容
			// run.setBold(true); // 加粗
			run.setFontSize(12);// 字体大小
			run.setFontFamily("宋体");
			
			run.addBreak();// 添加一个回车空行
		}
		return document;		
	}
	
	// 初始化XWPFDocument
	public static XWPFDocument initDoc(List<Map<String, Object>> list) throws Exception {
		XWPFDocument document = new XWPFDocument();
		// 设置边距
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		CTPageMar pageMar = sectPr.addNewPgMar();
		pageMar.setLeft(BigInteger.valueOf(720L));
		pageMar.setTop(BigInteger.valueOf(1440L));
		pageMar.setRight(BigInteger.valueOf(720L));
		pageMar.setBottom(BigInteger.valueOf(1440L));

		// ================插入图片======================
		XWPFParagraph pictures = document.createParagraph();
		pictures.setIndentationLeft(300);// 前面缩进300
		XWPFRun insertNewRun = pictures.createRun();
		int format = 0;
		String imgFile = "D:\\123.jpg";

		if (imgFile.endsWith(".emf"))
			format = XWPFDocument.PICTURE_TYPE_EMF;
		else if (imgFile.endsWith(".wmf"))
			format = XWPFDocument.PICTURE_TYPE_WMF;
		else if (imgFile.endsWith(".pict"))
			format = XWPFDocument.PICTURE_TYPE_PICT;
		else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg"))
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		else if (imgFile.endsWith(".png"))
			format = XWPFDocument.PICTURE_TYPE_PNG;
		else if (imgFile.endsWith(".dib"))
			format = XWPFDocument.PICTURE_TYPE_DIB;
		else if (imgFile.endsWith(".gif"))
			format = XWPFDocument.PICTURE_TYPE_GIF;
		else if (imgFile.endsWith(".tiff"))
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		else if (imgFile.endsWith(".eps"))
			format = XWPFDocument.PICTURE_TYPE_EPS;
		else if (imgFile.endsWith(".bmp"))
			format = XWPFDocument.PICTURE_TYPE_BMP;
		else if (imgFile.endsWith(".wpg"))
			format = XWPFDocument.PICTURE_TYPE_WPG;
		else {
			System.err.println("Unsupported picture: " + imgFile + ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");

		}
		insertNewRun.addPicture(new FileInputStream(imgFile), format, imgFile, Units.toEMU(350), Units.toEMU(150)); // 高150，宽350
		// insertNewRun.addBreak(BreakType.PAGE);//图片独占一页
		insertNewRun.addBreak();// 添加一个回车空行

		// ====================设置段落======================
		XWPFParagraph firstParagraph = document.createParagraph();
		firstParagraph.setAlignment(ParagraphAlignment.RIGHT);// 设置段落内容靠右
		firstParagraph.setIndentationRight(500);// 末尾缩进300

		XWPFRun run = firstParagraph.createRun();
		run.setText("测试段落   样式设置");
		run.setBold(true); // 加粗
		run.setFontSize(12);// 字体大小
		run.setFontFamily("华文中宋");

		run.addBreak();// 添加一个回车空行

		// ============添加标题====================
		XWPFParagraph titleParagraph = document.createParagraph();
		// 设置段落居中
		titleParagraph.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun titleParagraphRun = titleParagraph.createRun();
		titleParagraphRun.setText("测试使用poi导出word");
		titleParagraphRun.setBold(true);// 加粗
		titleParagraphRun.setFontSize(22);// 字体大小
		titleParagraphRun.setFontFamily("华文中宋");

		titleParagraphRun.addBreak();
		// ==============添加表格数据===========
		// 数据表格
		XWPFTable ComTable = document.createTable();// 默认创建一个一行一列的表格

		// ===========表格表头行===============
		XWPFTableRow titleRow = ComTable.getRow(0);// 创建的的一行一列的表格，获取第一行
		titleRow.setHeight(500);// 设置当前行行高
		titleStyle(titleRow);// 设置表头行样式和内容

		// 将数据封装到表格中
		for (Map<String, Object> m : list) {
			// =============合并行----合并单位名称======================
			XWPFTableRow DWMCRow = ComTable.createRow();// 创建表头行下新的一行，列数是按照表头行的列数创建的
			DWMCRow.setHeight(600);// 设置行高
			// 获取当前行的列数，原理第一行有几列，下面创建的新行就有几列，如果创建表格时指定了，就按指定的数创建多少列
			List<XWPFTableCell> DWMCCellList = DWMCRow.getTableCells();
			for (int i = 0; i < DWMCCellList.size(); i++) {// 经过这个循环后就将这一行的所有列合并成一列了
				// 对单元格进行合并的时候,要标志单元格是否为起点,或者是否为继续合并
				if (i == 0)
					DWMCCellList.get(i).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);// 这是起点
				else
					DWMCCellList.get(i).getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);// 继续合并
			}
			XWPFTableCell DWMCCell = DWMCCellList.get(0);// 当前行合并成一列了，获取这一列
			TitleCelldata1(DWMCCell, m.get("dwmc") == null ? "无" : m.get("dwmc").toString());// 给列添加样式和数据

			// =================================遍历具体的数据===============================
			List<Map<String, Object>> list1 = (List<Map<String, Object>>) m.get("data");
			int i = 0;
			for (Map<String, Object> dm : list1) {
				XWPFTableRow dataRow = ComTable.createRow();// 创建一个新的行
				dataRow.setHeight(500);// 设置行高
				List<XWPFTableCell> DataCellList = dataRow.getTableCells();// 获取创建行的所有的列
				XWPFTableCell tc0 = DataCellList.get(0);// 获取创建行的第一列
				CelldataCss(tc0, "800");// 给第一列序号添加样式
				tc0.setText(++i + "");// 给第一列添加值
				for (int ci = 1; ci < DataCellList.size(); ci++) {// 循环给剩下的列赋值
					XWPFTableCell tc = DataCellList.get(ci);
					// CelldataCss(tc,"1300");//给列添加样式
					CellData(dm, tc, ci);// 填充数据
				}
			}
		}
		return document;
	}

	// 设置表头行样式和内容
	public static void titleStyle(XWPFTableRow titleRow) {
		XWPFTableCell xhCell = titleRow.getCell(0);// 创建的表格没有指定几行几列，就默认是一行一列，这里获取创建的第一列
		xhCell.setText("");
		CelldataCss(xhCell, "800");
		XWPFTableCell bmCell = titleRow.addNewTableCell();// 在当前行继续创建新列

		TitleCelldata(bmCell, "部门");

		XWPFTableCell xmCell = titleRow.addNewTableCell();
		TitleCelldata(xmCell, "姓名");

		XWPFTableCell zwCell = titleRow.addNewTableCell();
		TitleCelldata(zwCell, "职务");

		XWPFTableCell bgdhCell = titleRow.addNewTableCell();
		TitleCelldata(bgdhCell, "办公电话");

		XWPFTableCell sjhmCell = titleRow.addNewTableCell();
		TitleCelldata(sjhmCell, "手机号码");

		XWPFTableCell yxCell = titleRow.addNewTableCell();
		TitleCelldata(yxCell, "邮箱");

	}

	// 设置数据列样式
	public static void CelldataCss(XWPFTableCell cell, String width) {
		/** 设置水平居中 */
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);// 上下居中
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);// 左右居中
		CTTblWidth tblWidth = ctPr.isSetTcW() ? ctPr.getTcW() : ctPr.addNewTcW();
		tblWidth.setW(new BigInteger(width));// 设置列宽度
		tblWidth.setType(STTblWidth.DXA);

	}

	// 设置内容信息样式
	public static void dataCelldata(XWPFTableCell cell, String txt, String tname) {
		// 给当前列中添加段落，就是给列添加内容
		XWPFParagraph p = cell.getParagraphs().get(0);
		XWPFRun headRun0 = p.createRun();
		headRun0.setText(txt);// 设置内容
		headRun0.setFontSize(11);// 设置大小
		// headRun0.setBold(true);//是否粗体
		headRun0.setFontFamily("仿宋_GB2312");
		// 给列中的内容设置样式
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);// 上下居中
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);// 左右居中
		CTTblWidth tblWidth = ctPr.isSetTcW() ? ctPr.getTcW() : ctPr.addNewTcW();
		if ("zw".equals(tname)) {
			tblWidth.setW(new BigInteger("2050"));// 设置列宽度
		} else {
			tblWidth.setW(new BigInteger("1600"));// 设置列宽度
		}
		tblWidth.setType(STTblWidth.DXA);
	}

	// 设置表头和单位信息样式
	public static void TitleCelldata(XWPFTableCell cell, String txt) {
		// 给当前列中添加段落，就是给列添加内容
		XWPFParagraph p = cell.getParagraphs().get(0);
		XWPFRun headRun0 = p.createRun();
		headRun0.setText(txt);// 设置内容
		headRun0.setFontSize(12);// 设置大小
		headRun0.setBold(true);// 是否粗体
		headRun0.setFontFamily("华文中宋");
		// 给列中的内容设置样式
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);// 上下居中
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);// 左右居中
		CTTblWidth tblWidth = ctPr.isSetTcW() ? ctPr.getTcW() : ctPr.addNewTcW();
		if ("职务".equals(txt)) {
			tblWidth.setW(new BigInteger("2050"));// 设置列宽度
		} else {
			tblWidth.setW(new BigInteger("1600"));// 设置列宽度
		}

		tblWidth.setType(STTblWidth.DXA);
	}

	// 设置表头和单位信息样式
	public static void TitleCelldata1(XWPFTableCell cell, String txt) {
		// 给当前列中添加段落，就是给列添加内容
		XWPFParagraph p = cell.getParagraphs().get(0);
		XWPFRun headRun0 = p.createRun();
		headRun0.setText(txt);// 设置内容
		headRun0.setFontSize(12);// 设置大小
		headRun0.setBold(true);// 是否粗体
		headRun0.setFontFamily("华文中宋");
		// 给列中的内容设置样式
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);// 上下居中
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);// 左右居中
		CTTblWidth tblWidth = ctPr.isSetTcW() ? ctPr.getTcW() : ctPr.addNewTcW();
		tblWidth.setW(new BigInteger(1600 * 6 + 800 + 450 + ""));// 设置列宽度
		tblWidth.setType(STTblWidth.DXA);
		cell.setColor("DCDCDC");
	}

	// 将数据赋值到每一列上
	public static void CellData(Map<String, Object> m, XWPFTableCell cell, Integer ic) {
		switch (ic) {
		case 1:
			dataCelldata(cell, m.get("bm") == null ? "无" : m.get("bm").toString(), "bm");
			break;
		case 2:
			dataCelldata(cell, m.get("mc") == null ? "无" : m.get("mc").toString(), "mc");
			break;
		case 3:
			dataCelldata(cell, m.get("zw") == null ? "无" : m.get("zw").toString(), "zw");
			break;
		case 4:
			dataCelldata(cell, m.get("bgdh") == null ? "无" : m.get("bgdh").toString(), "bgdh");
			break;
		case 5:
			dataCelldata(cell, m.get("sjhm") == null ? "无" : m.get("sjhm").toString(), "sjhm");
			break;
		case 6:
			dataCelldata(cell, m.get("yx") == null ? "无" : m.get("yx").toString(), "yx");
			break;
		}
	}

	// 手动构造数据
	public static List<Map<String, Object>> data() {
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();

		Map<String, Object> m1 = new HashMap<String, Object>();
		m1.put("dwmc", "测试部门1");
		List<Map<String, Object>> list1 = new ArrayList<Map<String, Object>>();
		for (int i = 0; i < 5; i++) {
			Map<String, Object> d1 = new HashMap<String, Object>();
			d1.put("bm", "部门" + i);// 部门
			d1.put("mc", "名称" + i);// 名称
			d1.put("zw", "职务" + i);// 职务
			d1.put("bgdh", "办公电话" + i);// 办公电话
			d1.put("sjhm", "手机号码" + i);// 手机号码
			d1.put("yx", "邮箱" + i);// 邮箱
			d1.put("xh", i + 1);// 序号
			list1.add(d1);
		}
		m1.put("data", list1);

		Map<String, Object> m2 = new HashMap<String, Object>();
		m2.put("dwmc", "测试部门2");
		List<Map<String, Object>> list2 = new ArrayList<Map<String, Object>>();
		for (int i = 0; i < 3; i++) {
			Map<String, Object> d1 = new HashMap<String, Object>();
			d1.put("bm", "部门" + i);// 部门
			d1.put("mc", "名称" + i);// 名称
			d1.put("zw", "职务" + i);// 职务
			d1.put("bgdh", "办反反复复烦烦烦烦烦烦烦烦烦烦烦烦烦烦烦吼吼吼吼吼吼吼吼吼吼吼吼吼吼吼吼公电话" + i);// 办公电话
			d1.put("sjhm", "手机号码" + i);// 手机号码
			d1.put("yx", "邮箱" + i);// 邮箱
			d1.put("xh", i + 1);// 序号
			list2.add(d1);
		}
		m2.put("data", list2);

		Map<String, Object> m3 = new HashMap<String, Object>();
		m3.put("dwmc", "测试部门3");
		List<Map<String, Object>> list3 = new ArrayList<Map<String, Object>>();
		for (int i = 0; i < 2; i++) {
			Map<String, Object> d1 = new HashMap<String, Object>();
			d1.put("bm", "部门" + i);// 部门
			d1.put("mc", "名称" + i);// 名称
			d1.put("zw", "职务" + i);// 职务
			d1.put("bgdh", "办公电话" + i);// 办公电话
			d1.put("sjhm", "手机号码" + i);// 手机号码
			d1.put("yx", "邮箱" + i);// 邮箱
			d1.put("xh", i + 1);// 序号
			list3.add(d1);
		}
		m3.put("data", list3);
		list.add(m1);
		list.add(m2);
		list.add(m3);
		return list;
	}

	/**
	 * insert Picture
	 * 
	 * @param document
	 * @param filePath
	 * @param inline
	 * @param width
	 * @param height
	 * @throws InvalidFormatException
	 * @throws FileNotFoundException
	 */
	private static void insertPicture(XWPFDocument document, String filePath, CTInline inline, int width, int height)
			throws InvalidFormatException, FileNotFoundException {
		document.addPictureData(new FileInputStream(filePath), XWPFDocument.PICTURE_TYPE_PNG);
		int id = document.getAllPictures().size() - 1;
		final int EMU = 9525;
		width *= EMU;
		height *= EMU;
		// String blipId =
		// document.getAllPictures().get(id).getPackageRelationship().getId();
		// blipId
		// =document.addPictureData(fastDFSClient.download(src),Document.PICTURE_TYPE_PNG);

		// String blipId =
		// document.getAllPictures().get(id).getPackageRelationship().getId();
		String blipId = document.getAllPackagePictures().get(id)
				.getRelationId(document.getAllPackagePictures().get(id).getParent());
		String picXml = getPicXml(blipId, width, height);
		XmlToken xmlToken = null;
		try {
			xmlToken = XmlToken.Factory.parse(picXml);
		} catch (XmlException xe) {
			xe.printStackTrace();
		}
		inline.set(xmlToken);
		inline.setDistT(0);
		inline.setDistB(0);
		inline.setDistL(0);
		inline.setDistR(0);
		CTPositiveSize2D extent = inline.addNewExtent();
		extent.setCx(width);
		extent.setCy(height);
		CTNonVisualDrawingProps docPr = inline.addNewDocPr();
		docPr.setId(id);
		docPr.setName("IMG_" + id);
		docPr.setDescr("IMG_" + id);
	}

	/**
	 * get the xml of the picture
	 * 
	 * @param blipId
	 * @param width
	 * @param height
	 * @return
	 */
	private static String getPicXml(String blipId, int width, int height) {
		String picXml = "" + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
				+ "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
				+ "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
				+ "         <pic:nvPicPr>" + "            <pic:cNvPr id=\"" + 0 + "\" name=\"Generated\"/>"
				+ "            <pic:cNvPicPr/>" + "         </pic:nvPicPr>" + "         <pic:blipFill>"
				+ "            <a:blip r:embed=\"" + blipId
				+ "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
				+ "            <a:stretch>" + "               <a:fillRect/>" + "            </a:stretch>"
				+ "         </pic:blipFill>" + "         <pic:spPr>" + "            <a:xfrm>"
				+ "               <a:off x=\"0\" y=\"0\"/>" + "               <a:ext cx=\"" + width + "\" cy=\""
				+ height + "\"/>" + "            </a:xfrm>" + "            <a:prstGeom prst=\"rect\">"
				+ "               <a:avLst/>" + "            </a:prstGeom>" + "         </pic:spPr>"
				+ "      </pic:pic>" + "   </a:graphicData>" + "</a:graphic>";
		return picXml;
	}
	
	/**
	 * 写入文件
	 * @param data
	 * @param out
	 * @throws Exception
	 */
	public static void exportWord(List<Map<String, Object>> data, OutputStream out) throws Exception {
        // 导出word文件
		XWPFDocument document = createParagraphDoc(data);
        try {
            document.write(out);
        } finally {
            if(out != null){
                try {
                	out.close();
                } catch (Exception e) {
                    logger.error(e.getMessage(),e);
                }
            }
            if(document != null){
            	try {
            		document.close();
            	} catch (Exception e) {
            		logger.error(e.getMessage(),e);
            	}
            }
        }
    }
	
	/**
	 * 导出word文档
	 * @param response
	 * @param fileName
	 * @param data
	 * @throws Exception
	 */
	public static void exportWord(HttpServletResponse response, String fileName, List<Map<String, Object>> data) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/vnd.ms-excel");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(fileName, "utf-8"));
        // 开始写入文件
        exportWord(data, response.getOutputStream());
    }
	
	public static void main(String[] args) throws Exception {
		XWPFDocument document = initDoc(data());
		FileOutputStream out = new FileOutputStream("d:\\部门通讯录.docx");
		document.write(out);
		out.close();
		System.out.println("导出成功!!!!!!!!");
	}

}
