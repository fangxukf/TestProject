import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
public class XwpfTest{
    public static void main(String[] args)throws Exception {
        InputStream is = new FileInputStream("C:\\Users\\Administrator\\Desktop\\abc.doc");
        @SuppressWarnings("resource")
        XWPFDocument doc = new XWPFDocument(is);
        List<XWPFParagraph> paras = doc.getParagraphs();//将得到包含段落列表
        System.out.println("all data :" + paras.size());
        for(XWPFParagraph para : paras) {
            //当前段落的属性  
            //CTPPr pr = para.getCTP().getPPr(); 
            //System.out.println(para.getText());
            List<XWPFRun> runsLists = para.getRuns();//获取段楼中的句列表

            for(XWPFRun runsList : runsLists ){
                String c = runsList.getColor();//获取句的字体颜色
                float f = runsList.getFontSize();//获取句中字的大小
                String s = runsList.getText(0);//获取文本内容

                if(s != null) // 如果读取为非空，则对其进行判断
                {
                    if(s.contains("摘要"))// 识别摘要
                    {
                        System.out.println("right!");
                        runsList.setBold(true);
                    }
                    if(s.equals("摘要：")){
                        System.out.println("ddddddddddddddddddd");
                    }

                    if(s.contains("第一章")){
                        if(f != 16){
                            System.out.println("一级标题格式不是三号字体！");
                        }
                        System.out.println("一级标题！！！！！");
                    }
                }

                System.out.println("color:" + c);
                System.out.println("size:" + f);
                System.out.print("text:" + s);
                if(s != null){
                    System.out.print(",the length of string is " + s.length());
                }
                System.out.println("-----");
            }
        }
    }
}