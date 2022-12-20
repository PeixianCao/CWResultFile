import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import jxl.Cell;
import jxl.JXLException;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class CWOutputFile {
    /**
     * wOutputFile 作用：把内容写入到Excel文件中。
     * wOutputFile写结果文件   wOutputFile(文件路径、用例编号、用例标题、预期结果、实际结果、测试结果)
     * @throws IOException
     * @throws BiffException
     * @throws WriteException
     */
    public void wOutputFile(String filepath,String caseNo,String testPoint,String testData,String preResult,String fresult) throws BiffException, IOException, WriteException{
        System.out.println("pull");
        File output=new File(filepath);
        String result = "";
        InputStream instream = new FileInputStream(filepath);
        Workbook readwb = Workbook.getWorkbook(instream);
        WritableWorkbook wbook = Workbook.createWorkbook(output, readwb);  //根据文件创建一个操作对象
        WritableSheet readsheet = wbook.getSheet(0);  //定位到文件的第一个sheet页签
        int rsRows = readsheet.getRows();   //获取sheet页签的总行数
        //获取sheet表中所包含的总行数

        /******************设置字体样式***************************/
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"),10,WritableFont.NO_BOLD);
        WritableCellFormat wcf = new WritableCellFormat(font);
        /****************************************************/

        Cell cell = readsheet.getCell(0,rsRows);  //获取sheet页的单元格
        if(cell.getContents().equals("")){
            Label labetest1 = new Label(0,rsRows,caseNo);   //第一列：用例编号
            Label labetest2 = new Label(1,rsRows,testPoint);//第二列：用例标题
            Label labetest3 = new Label(2,rsRows,testData); //第三列：测试数据
            Label labetest4 = new Label(3,rsRows,preResult);//第四列：预期结果
            Label labetest5 = new Label(4,rsRows,fresult); //第五列：实际结果
            if(preResult.equals(fresult)){ // 判断两个值同时相等才会显示通过
                result = "通过"; //预期结果和实际结果相同，测试通过
                wcf.setBackground(Colour.BRIGHT_GREEN);  //通过案例标注绿色
            }
            else{
                result = "不通过"; //预期结果和实际结果不相同，测试不通过
                wcf.setBackground(Colour.RED);  // 不通过案例标注红色
            }
            Label labetest6 = new Label(5,rsRows,result,wcf);//第六列：测试结果
            readsheet.addCell(labetest1);
            readsheet.addCell(labetest2);
            readsheet.addCell(labetest3);
            readsheet.addCell(labetest4);
            readsheet.addCell(labetest5);
            readsheet.addCell(labetest6);
        }
        wbook.write();
        wbook.close();
    }

    /**
     * cOutputFile 作用：创建Excel文件,
     * tradeType为文件名称前缀，
     * 返回结果：文件路径，作为wOutputFile写入结果文件的入参
     * @throws IOException
     * @throws WriteException
     * */
    public String cOutputFile(String tradeType) throws IOException, WriteException{
        String temp_str = "";
        Date dt = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        temp_str = sdf.format(dt); //获取时间戳
        // 相对路径默认为 JMeter_home\bin 目录
        // 以时间戳命名结果文件，确保唯一
        // 生成文件路径
        String filepath = "D:\\\\"+tradeType+"_output_" + "_" + temp_str + ".xls";
        File output = new File(filepath);
        if(!output.isFile()){
            // 如果指定的文件不存在，创建新该文件
            output.createNewFile();
            // 写文件
            // 新建一个writeBook，在新建一个sheet
            WritableWorkbook writeBook = Workbook.createWorkbook(output);
            //命名sheet // createsheet(sheet名称，第几个sheet)
            WritableSheet sheet = writeBook.createSheet("输出结果", 0);
            //设置首行字体为宋体，11号，加粗
            WritableFont headfont = new WritableFont(WritableFont.createFont("宋体"),11,WritableFont.BOLD);
            WritableCellFormat headwcf = new WritableCellFormat(headfont);
            headwcf.setBackground(Colour.GRAY_25); // 灰色颜色
            // 设置列宽度setcolumnview(列号，宽度)
            sheet.setColumnView(0, 11); //设置列宽
            sheet.setColumnView(1, 20);
            sheet.setColumnView(2, 40);
            sheet.setColumnView(3, 10);
            sheet.setColumnView(4, 10);
            sheet.setColumnView(5, 10); //如果需要再新增列，这里往下添加
            headwcf.setAlignment(Alignment.CENTRE); //设置文字居中对齐方式；//文字居中
            headwcf.setVerticalAlignment(VerticalAlignment.CENTRE); // 设置垂直居中；
            Label labe00 = new Label(0,0,"用例编号",headwcf); //写入内容：Label(列号，行号，内容)
            Label labe10 = new Label(1,0,"用例标题",headwcf);
            Label labe20 = new Label(2,0,"测试数据",headwcf);
            Label labe30 = new Label(3,0,"预期结果",headwcf);
            Label labe40 = new Label(4,0,"实际结果",headwcf);
            Label labe50 = new Label(5,0,"执行结果",headwcf); //往下添加
            sheet.addCell(labe00);
            sheet.addCell(labe10);
            sheet.addCell(labe20);
            sheet.addCell(labe30);
            sheet.addCell(labe40);
            sheet.addCell(labe50);//往下添加
            writeBook.write();
            writeBook.close();
        }
        return filepath;
    }
}
