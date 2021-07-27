package com.ssm.test;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author zf
 * @date 2021/7/26 9:10
 */
public class WordMatchUtil {
    //手机号正则
    private static String REX_PHONE= "(1([358][0-9]|4[579]|66|7[0135678]|9[89])[0-9]{8})(?!(\\.))(?!(\\d+))";
    //身份证正则
    private static String REX_IDCARD= "([1-9]\\d{5}(18|19|20)\\d{2}((0[1-9])|(1[0-2]))(([0-2][1-9])|10|20|30|31)\\d{3}[0-9Xx])(?!(\\.))(?!(\\d{1,20}))";

    /**
     * 2007版本word转换成html
     * @throws IOException
     */
    public static boolean Word2007ToHtml(String file){
        InputStream inputStream = null;
        try{
            /* File f = new File(file);
           if (!f.exists()) {
                System.out.println("Sorry File does not Exists!");
            } else {*/
                if (file.endsWith(".docx") || file.endsWith(".DOCX")) {

                    //加载word文档生成 XWPFDocument对象
                    //InputStream in = new FileInputStream(f);
                    inputStream = HttpFIleUtil.inputStream(file);
                    XWPFDocument document = new XWPFDocument(inputStream);

                    XHTMLOptions options = XHTMLOptions.create();
                    //options.setExtractor(new FileImageExtractor(imageFolderFile));
                    options.setIgnoreStylesIfUnused(false);
                    options.setFragment(true);

                    //使用字符数组流获取解析的内容
                    ByteArrayOutputStream baos = new ByteArrayOutputStream();
                    XHTMLConverter.getInstance().convert(document, baos, options);
                    String content = baos.toString();
                    content = contentFilter(content);
                    //匹配手机号和身份证
                    boolean phone = isPhone(content);
                    boolean idCard = isIDCard(content);
                    baos.close();
                    if(phone || idCard){
                        return true;
                    }else {
                        return false;
                    }
                /*if(phone && idCard){
                    return "手机号和身份证号";
                }else if (phone && !idCard){
                    return "手机号";
                }else if(!phone && idCard){
                    return "身份证号";
                }else {
                    return null;
                }*/
                } else {
                    System.out.println("Enter only MS Office 2007+ files");
                }
            //}

        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (null != inputStream) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        //return null;
        return false;
    }

    /**
     * /**
     * 2003版本word转换成html
     * @throws IOException
     * @throws TransformerException
     * @throws ParserConfigurationException
     */
    public static boolean Word2003ToHtml(String file) {
        InputStream inputStream = null;
        try {
            //File f = new File(file);
            if (file.endsWith(".doc") || file.endsWith(".DOC")) {
                inputStream = HttpFIleUtil.inputStream(file);
                HWPFDocument wordDocument = new HWPFDocument(inputStream);
                WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                        DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());

                // 解析word文档
                wordToHtmlConverter.processDocument(wordDocument);
                Document htmlDocument = wordToHtmlConverter.getDocument();

                // 使用字符数组流获取解析的内容
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                DOMSource domSource = new DOMSource(htmlDocument);
                StreamResult streamResult = new StreamResult(baos);

                TransformerFactory factory = TransformerFactory.newInstance();
                Transformer serializer = factory.newTransformer();
                serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
                serializer.setOutputProperty(OutputKeys.INDENT, "yes");
                serializer.setOutputProperty(OutputKeys.METHOD, "html");
                serializer.transform(domSource, streamResult);

                String content = new String(baos.toByteArray());
                content = contentFilter(content);
                //匹配手机号和身份证
                boolean phone = isPhone(content);
                boolean idCard = isIDCard(content);
                baos.close();
                if(phone || idCard){
                    return true;
                }else {
                    return false;
                }
            /*if(phone && idCard){
                return "手机号和身份证号";
            }else if (phone && !idCard){
                return "手机号";
            }else if(!phone && idCard){
                return "身份证号";
            }else {
                return null;
            }*/
            } else {
                System.out.println("Enter only MS Office 2003 files");

            }
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (null != inputStream) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        //return null;
        return false;
    }

    /**
     * 内容过滤
     * @param content
     * @return
     */
    public static String contentFilter(String content){
        content = HtmlUtil.getTextFromHtml(content);//过滤html标签
        content = content.replaceAll("&ensp;", "");
        content = ProofreadUtil.strFilter4JsonConflict(content);//过滤Json特殊字符
        return content;
    }

    /**
     * 验证手机号
     * @param str
     * @return
     */
    public static boolean isPhone(final String str) {
        Pattern p = null;
        Matcher m = null;
        boolean b = false;
        p = Pattern.compile(REX_PHONE);
        m = p.matcher(str);
        b = m.find();
        return b;
    }

    /**
     * 验证身份证
     * @param str
     * @return
     */
    public static boolean isIDCard(final String str) {
        Pattern p = null;
        Matcher m = null;
        boolean b = false;
        p = Pattern.compile(REX_IDCARD);
        m = p.matcher(str);
        b = m.find();
        return b;
    }
}
