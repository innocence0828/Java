package com.test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComException;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class demo {
	// 声明一个word对象
	private ActiveXComponent objWord;
	// 声明四个word组件
	private Dispatch document;
	private Dispatch wordObject;
	// 选定的范围或插入点
	private static Dispatch selection = null;
	public demo() {

	}

	/**
	 * 打开word文挡
	 */
	public void open(String filename) {
		// 创建一个word对象
		objWord = new ActiveXComponent("Word.Application");
		// 为wordobject组件附值
		wordObject = (Dispatch) (objWord.getObject()); 
		// 生成一个只读方式的word文挡组件
		Dispatch.put(wordObject, "Visible", new Variant(false));
		// 获取文挡属性
		Dispatch documents = objWord.getProperty("Documents").toDispatch();
		// 打开激活文挡
		document = Dispatch.call(documents, "Open", filename).toDispatch();

		// 让文档显示最终状态,提高执行效率
		Dispatch actWin = Dispatch.get(objWord, "ActiveWindow").toDispatch();
		Dispatch view = Dispatch.get(actWin, "View").toDispatch();
		Dispatch.put(view, "ShowRevisionsAndComments", new Variant(false));
		Dispatch.put(view, "RevisionsView", new Variant(0));
		selection = objWord.getProperty("Selection").toDispatch();
	}

	/**
	 * 关闭文挡
	 */
	public void close() {
		Dispatch.call(document, "Close");
	}
	
	/**
	 * 把插入点移动到文件首位置
	 */
	public void moveStart() {
		if (selection == null)
			selection = Dispatch.get(document, "Selection").toDispatch();
		Dispatch.call(selection, "HomeKey", new Variant(6));
	}

	/**
	 * 复制模板文件为副本
	 * @param oldPath
	 * @param newPath
	 */
	public void Copy(String oldPath, String newPath) {
		InputStream inStream = null;
		FileOutputStream fs = null;
		try {
			int bytesum = 0;
			int byteread = 0;
			File oldfile = new File(oldPath);
			if (oldfile.exists()) {
				inStream = new FileInputStream(oldPath);
				fs = new FileOutputStream(newPath);
				byte[] buffer = new byte[1444];
				int length;
				while ((byteread = inStream.read(buffer)) != -1) {
					bytesum += byteread;
					fs.write(buffer, 0, byteread);
				}
			}
		} catch (Exception e) {
			System.out.println("error  ");
			e.printStackTrace();
		} finally {

			// 最后一定要关闭
			if (inStream != null) {
				try {
					inStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			if (fs != null) {
				try {
					fs.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * 从选定内容或插入点开始查找文本
	 * 
	 * @param selection
	 *            Dispatch 选定内容
	 * @param toFindText
	 *            String 要查找的文本
	 * @return boolean true-查找到并选中该文本，false-未查找到文本
	 */
	public boolean find(String toFindText) {
		// 从selection所在位置开始查询
		Dispatch find = objWord.call(selection, "Find").toDispatch();
		// 设置要查找的内容
		Dispatch.put(find, "Text", toFindText);
		// 向前查找
		Dispatch.put(find, "Forward", "True");
		// 设置格式
		Dispatch.put(find, "Format", "True");
		// 大小写匹配
		Dispatch.put(find, "MatchCase", "True");
		// 全字匹配
		Dispatch.put(find, "MatchWholeWord", "True");
		// 查找并选中
		return Dispatch.call(find, "Execute").getBoolean();
	}

	/**
	 *  全文替换,即文章中所有的字符串都被替换,去掉while则替换一次.
	 * @param oldtext
	 * @param newtext
	 */
	public void wordFindReplace(String oldtext, String newtext) {
		moveStart();
		while (find(oldtext)) {
			Dispatch.put(selection, "Text", newtext);
			Dispatch.call(selection, "MoveRight");
		}
	}
	
	/** 创建表格
	    * @param pos 位置
	    * @param cols 列数
	    * @param rows 行数
	    */
	   public Dispatch createTable(String tableTitle, int numCols, int numRows) {
		   Dispatch selection = Dispatch.get(objWord, "Selection").toDispatch(); // 输入内容需要的对象
		   Dispatch.call(selection, "TypeParagraph"); // 空一行段落
		   Dispatch.call(selection, "TypeText", tableTitle); // 写入标题内容 标题格行
		   Dispatch tables = Dispatch.get(document, "Tables").toDispatch();
	       Dispatch range = Dispatch.get(selection, "Range").toDispatch(); //当前光标位置或者选中的区域
	       Dispatch newTable = Dispatch.call(tables, "Add", range, new Variant(numRows), 
	    		   new Variant(numCols), new Variant(1)).toDispatch();
	       return newTable;
	   }
	   
		/**
		 * 在指定的单元格里填写数据
		 * @param tableIndex
		 * @param cellRowIdx
		 * @param cellColIdx
		 * @param txt
		 */
		public void putTxtToCell(Dispatch table, int cellRowIdx, int cellColIdx, String txt) {
			Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),
					new Variant(cellColIdx)).toDispatch();
			Dispatch.call(cell, "Select");
			Dispatch selection = Dispatch.get(objWord, "Selection").toDispatch(); // 输入内容需要的对象
			Dispatch.put(selection, "Text", txt);
		}
		
		/**
		 * 设置当前选定内容的字体
		 * @param boldSize
		 * @param italicSize
		 * @param underLineSize
		 *            下划线
		 * @param colorSize
		 *            字体颜色
		 * @param size
		 *            字体大小
		 * @param name
		 *            字体名称
		 */
		public void setFont(boolean  isBold,  boolean  isItalic,String colorSize, String size) {
			Dispatch font = Dispatch.get(selection, "Font").toDispatch();
			Dispatch.put(font, "Bold", isBold);
			Dispatch.put(font, "Italic", isItalic);
			Dispatch.put(font, "Color", colorSize);
			Dispatch.put(font, "Size", size);
		}
		
		/**
		* 像word中插入图片
		* @param imagePath 图片路径
		* @param width 宽度
		* @param height 高度
		*/
		public void insertImage(String imagePath,int width,int height)
		{
			Dispatch picture = Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(), "AddPicture", imagePath).toDispatch();
			Dispatch.put(picture, "Width", new Variant(width));
			Dispatch.put(picture, "Height", new Variant(height));
		}
		
		 /** 
	     * 合并单元格 
	     * @param tableIndex 
	     * @param fstCellRowIdx 
	     * @param fstCellColIdx 
	     * @param secCellRowIdx 
	     * @param secCellColIdx 
	     */ 
	    public void mergeCell1(Dispatch table, int fstCellRowIdx, int fstCellColIdx, 
	                    int secCellRowIdx, int secCellColIdx) { 
	            Dispatch fstCell = Dispatch.call(table, "Cell", 
	                            new Variant(fstCellRowIdx), new Variant(fstCellColIdx)) 
	                            .toDispatch(); 
	            Dispatch secCell = Dispatch.call(table, "Cell", 
	                            new Variant(secCellRowIdx), new Variant(secCellColIdx)) 
	                            .toDispatch(); 
	            Dispatch.call(fstCell, "Merge", secCell); 
	    } 
	    
	    /**
		 * 设置单元格被选中
		 * @param tableIndex
		 * @param cellRowIdx
		 * @param cellColIdx
		 */
	    public void setTableCellSelected(Dispatch table, int cellRowIdx, int cellColIdx) {
			Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),
					new Variant(cellColIdx)).toDispatch();
			Dispatch.call(cell, "Select");
		}

		/**
		 * 设置选定单元格的垂直对起方式, 请使用setTableCellSelected选中一个单元格
		 * @param align
		 * 0-顶端, 1-居中, 3-底端
		 */
		public void setCellVerticalAlign(int verticalAlign) {
			Dispatch cells = Dispatch.get(selection, "Cells").toDispatch();
			Dispatch.put(cells, "VerticalAlignment", new Variant(verticalAlign));
		}

	public static void main(String[] args) {
		try {
			demo jacTest = new demo();
			// 先复制模板 副本
			String fileFrom = "E:\\word\\demo.doc";
			String fileOut = "E:\\word\\demo_5.doc";
			jacTest.Copy(fileFrom, fileOut);
			jacTest.open("E:\\word\\demo_5.doc");
			
			jacTest.wordFindReplace("@:JHH@", "YS20161121-0001");
			jacTest.wordFindReplace("@:ZDR@", "张三");  //查找并替换文字
			String barPath = "E:\\word\\bar.png";
			
			Dispatch.call(selection, "TypeParagraph"); // 空一行段落
			Dispatch.call(selection, "TypeParagraph"); // 空一行段落
			
			jacTest.insertImage(barPath,405,230);   //插入图片
			
			Dispatch.call(selection, "TypeParagraph"); // 空一行段落
			Dispatch.call(selection, "TypeParagraph"); // 空一行段落
			
			Dispatch newTable = jacTest.createTable("表格统计",6,3);  //创建表格
			String[] arr = {"部门","类型","总数","已扫","未扫","盘盈"};   //设置表格列名
			String[] arr1 = {"部门1","类型1","总数1","已扫1","未扫1","盘盈1"};   //设置表格列名
			String[] arr2 = {"部门1","类型2","总数2","已扫2","未扫2","盘盈2"};   //设置表格列名

			for(int m=1;m<=arr.length;m++){   //填充单元格第一行
				jacTest.putTxtToCell(newTable,1,m,arr[m-1]); 
				jacTest.setFont(false,false,"0,0,0", "10");
			}
			
			for(int i=1;i<=arr1.length;i++){  //填充单元格第二行
				//jacTest.mergeCell1(newTable,2,1,3,1);   //合并第一列的第二行和第三行单元格
				jacTest.putTxtToCell(newTable,2,i,arr1[i-1]);   //填充单元格的内容
				jacTest.setFont(false,false,"0,0,0", "10");  //设置字体样式
				jacTest.setTableCellSelected(newTable,2,i);  //设置单元格居中，先选中单元格
				jacTest.setCellVerticalAlign(1);     //设置居中
			}
			
				for(int j=1;j<=arr2.length;j++){  //填充单元格第三行
				jacTest.putTxtToCell(newTable,3,j,arr2[j-1]);   //填充单元格的内容
			}
			jacTest.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}