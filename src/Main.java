///*
//import com.spire.doc.*;
//import com.spire.doc.documents.paragraph;
//
//import javax.swing.text.document;
//import java.io.bufferedwriter;
//import java.io.file;
//import java.io.filewriter;
//import java.io.ioexception;
//
//public class gettitle {
//    public static void main(string[] args)throws ioexception {
//        //加载word测试文档
//        document doc = new document();
//        doc.loadfromfile("input.docx");
//
//        //保存标题内容到.txt文档
//        file file = new file("gettitle.txt");
//        if (file.exists())
//        {
//            file.delete();
//        }
//        file.createnewfile();
//        filewriter fw = new filewriter(file, true);
//        bufferedwriter bw = new bufferedwriter(fw);
//
//        //遍历section
//        for (int i = 0; i < doc.getsections().getcount(); i++)
//        {
//            section section = doc.getsections().get(i);
//            //遍历paragraph
//            for (int j = 0; j < section.getparagraphs().getcount(); j++)
//            {
//                paragraph paragraph = section.getparagraphs().get(j);
//
//                //获取标题
//                if ( paragraph.getstylename().matches("1"))//段落为“标题1”的内容
//                {
//                    //获取段落标题内容
//                    string text = paragraph.gettext();
//
//                    //写入文本到txt文档
//                    bw.write("标题1: "+ text + "\r");
//                }
//                //获取标题
//                if ( paragraph.getstylename().matches("2"))//段落为“标题2”的内容
//                {
//                    //获取段落标题内容
//                    string text = paragraph.gettext();
//
//                    //写入文本到txt文档
//                    bw.write("标题2: " + text + "\r");
//                }
//                //获取标题
//                if ( paragraph.getstylename().matches("3"))//段落为“标题3”的内容
//                {
//                    //获取段落标题内容
//                    string text = paragraph.gettext();
//
//                    //写入文本到txt文档
//                    bw.write("标题3: " + text+"\r");
//                }
//                //获取标题
//                if ( paragraph.getstylename().matches("4"))//段落为“标题4”的内容
//                {
//                    //获取段落标题内容
//                    string text = paragraph.gettext();
//
//                    //写入文本到txt文档
//                    bw.write("标题4: " + text+"\r");
//                }
//
//                bw.write("\n");
//            }
//
//        }
//        bw.flush();
//        bw.close();
//        fw.close();
//    }
//}*/
