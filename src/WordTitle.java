import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.ParagraphProperties;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * @author hp
 * 获取doc⽂档的标题
 */
public class WordTitle {
    public static void main(String[] args) throws Exception {
        String filePath = "C:\\Users\\Administrator\\Desktop\\gzh";
        //FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\abc.xls");
        //printWord(filePath);
        //getWordTitles(filePath);
        File dir = new File(filePath);
        File[] fileList = dir.listFiles();
        for(File file : fileList) {
            FileOutputStream fileOut2 = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\out\\"+"123"+".xls");
            File[] fileList2 = file.listFiles();
            String name = file.getName();
           readAndWriterTest3(fileList2[0].getAbsolutePath(),fileOut2,name);
            fileOut2.close();
        }
    }

    public static void printWord(String filePath) throws IOException {

        InputStream is = new FileInputStream(filePath);
        HWPFDocument doc = new HWPFDocument(is);

        Range r = doc.getRange();// ⽂档范围
        System.out.println(r.numParagraphs());
        for (int i = 0; i < r.numParagraphs(); i++) {
            Paragraph p = r.getParagraph(i);// 获取段落
            int numStyles = doc.getStyleSheet().numStyles();
            int styleIndex = p.getStyleIndex();
            if (numStyles > styleIndex) {

                StyleSheet style_sheet = doc.getStyleSheet();

                StyleDescription style = style_sheet.getStyleDescription(styleIndex);
                ParagraphProperties style1 = style_sheet.getParagraphStyle(styleIndex);

                String styleName = style.getName();// 获取每个段落样式名称
                //System.out.println(style_sheet);
                System.out.println(styleName);
                // 获取⾃⼰理想样式的段落⽂本信息
                //String styleLoving = "标题";
                String text = p.text();// 段落⽂本
                //if (styleName != null && styleName.contains(styleLoving)) {
                if (styleName.equals("标题")) {
                    System.out.println(text);
                }
            }
        }
        //doc.close();
    }

    public static List<String> getWordTitles(String path) throws IOException{
        File file = new File(path);
        FileInputStream is = new FileInputStream(file);
        List<String> list = new ArrayList<String>();
        XWPFDocument doc = new XWPFDocument(is);
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (XWPFParagraph graph : paras) {
            String text = graph.getParagraphText();
            String style = graph.getStyle();
            System.out.println(style);
            if ("1".equals(style)) {
                System.out.println(text+"--["+style+"]");
            }else if ("2".equals(style)) {
                System.out.println(text+"--["+style+"]");
            }else if ("3".equals(style)) {
                System.out.println(text+"--["+style+"]");
            }else{
                continue;
            }
            list.add(text);
        }
        return list;
    }

    public static void readAndWriterTest3(String path,FileOutputStream fileOut,String sheetname) throws IOException {
        List<String> list = new ArrayList<>();
        List<String> list2 = new ArrayList<>();
        try {
            String basePath = path;
            File dir = new File(basePath);
            File[] fileList = dir.listFiles();
            for(File file : fileList) {
                FileInputStream fis = new FileInputStream(file);
                HWPFDocument doc = new HWPFDocument(fis);
                Range r = doc.getRange();// 文档范围
                for (int i = 0; i < r.numParagraphs(); i++) {
                    Paragraph p = r.getParagraph(i);// 获取段落
                    int numStyles = doc.getStyleSheet().numStyles();
                    int styleIndex = p.getStyleIndex();
                    //System.out.println(numStyles+" "+styleIndex);

                    if (numStyles > styleIndex) {
                        StyleSheet style_sheet = doc.getStyleSheet();
                        StyleDescription style = style_sheet.getStyleDescription(styleIndex);
                        ParagraphProperties style1 = style_sheet.getParagraphStyle(styleIndex);

                        String styleName = style.getName();// 获取每个段落样式名称
                        //System.out.println(style_sheet);
                        //System.out.println(styleName);
                        // 获取自己理想样式的段落文本信息
                        //String styleLoving = "标题";
                        String text = p.text();// 段落文本
                        //if (styleName != null && styleName.contains(styleLoving)) {
                        if (text.contains(".")&&!Pattern.matches("[0-9].*", text)) {
                            //String text = p.text();// 段落文本
                            //if (!text.contains("，") && !text.contains("；") && !text.contains("。") && !text.contains("") && !text.contains("20") && !text.contains("https")) {
                            if (!text.contains("；") && !text.contains("。") && !text.contains("http")) {
                                //System.out.println(text);
                                list.add(text);
                                list2.add(file.getName().substring(0,10));
                            }
                        }
                    }
                }
                doc.close();
                fis.close();
            }
            writeToExcel(list,list2,fileOut,sheetname);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void writeToExcel(List<String> list,List<String> list2,FileOutputStream fileOut,String sheetname) throws IOException {
       /* if (CollectionUtils.isEmpty(list)) {
            return;
        }*/

        Workbook wb = new HSSFWorkbook();
        int oneSheetHeadRowNum = 1; //  头部非内容行数（如标题等）
        int oneSheetMaxRowNum = 65536;  //  excel 一个 sheet 最多支持行
        int oneSheetContentAvailableRowNum = oneSheetMaxRowNum - oneSheetHeadRowNum;
        int len = list.size();
        int needSheetSize = len / (oneSheetContentAvailableRowNum);
        int j = 0;
        for (int i = 0; i <= needSheetSize; i++) {
            Sheet sheet = wb.createSheet(sheetname);

            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("日期");
            row.createCell(1).setCellValue("消息");
            row.createCell(2).setCellValue("类型");
            row.createCell(3).setCellValue("地区");

            for (int k = 0; j < list.size() && k < oneSheetContentAvailableRowNum; k++, j++) {
                String string = list.get(j);
                String string1 = list2.get(j);
                Row contentRow = sheet.createRow(k + 1);
                contentRow.createCell(0).setCellValue(string1);
                contentRow.createCell(1).setCellValue(string);
                //contentRow.createCell(2).setCellValue(string);
            }
        }
        wb.write(fileOut);
    }
}