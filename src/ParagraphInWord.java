package com.silin;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ParagraphInWord {
    public static void main(String[] args) throws Exception {

        // new 一个空白文档
        XWPFDocument document = new XWPFDocument();

        // 在文件系统中编写文档，命名为：createparagraph.docx
        FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Administrator\\Desktop\\createparagraph.docx"));

        // 创建段落
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(
                "其实，人在脆弱时都想有份停靠，心在无助时都想要个依靠。若有人能解读忧伤，陪伴彷徨，看穿逞强。又何必死撑着非要坚强。小时候脸皮薄，别人两句硬话就要哭出泪来。后来慢慢面子坚硬起来，各种冷嘲热讽可以假装听不见。其实大多数时候，我们没资格优雅的活着。什么时候能厚着脸皮去生活了，你就真懂了生活。你要逼自己优秀，然后骄傲的生活，余生还长，何必慌张，以后的你，会为自己所做的努力，而感到庆幸，别在最好的年纪选择了安逸。新的一天，加油。");

        document.write(out);
        out.close();
        System.out.println("createparagraph.docx written successfully");
    }
}