package com.ssm.test;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

/**
 * @author zf
 * @date 2021/7/26 8:33
 */
@Controller
@RequestMapping("/test")
public class TestController  {

    /**
     * 判断附件内是否存在手机号和身份证号
     * @param files
     * @return
     * C:/work_space/test/测试.docx
     */
    @RequestMapping("isCheck")
    @ResponseBody
    public Object isCheck(@RequestParam(value = "files[]") String[] files) {
        boolean isMatch = false;
        if(files.length>0){
            for (int i = 0; i < files.length; i++) {
                if(files[i].indexOf("xls") >-1 ||files[i].indexOf("XLS") >-1){
                    isMatch= ExcelMatchUtil.ExcelToHtml(files[i]);
                    System.out.println("xls||xlsx");
                }else if (files[i].endsWith(".docx") || files[i].endsWith(".DOCX")) {
                    isMatch = WordMatchUtil.Word2007ToHtml(files[i]);
                    System.out.println("docx");
                }else if (files[i].endsWith(".doc") || files[i].endsWith(".DOC")) {
                    isMatch = WordMatchUtil.Word2003ToHtml(files[i]);
                    System.out.println("doc");
                }
                if (isMatch){
                    System.out.println(isMatch+"==="+files[i]);
                    return isMatch;
                }
            }
        }
        return null;
    }


}
