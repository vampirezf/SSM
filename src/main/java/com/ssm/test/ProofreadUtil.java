package com.ssm.test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName: ProofreadUtil
 * @Description: 校对工具类
 * @Author: zhushouyong
 * @CreateDate: 2019-7-31 17:27
 * @Version: 1.0
 */
public class ProofreadUtil {

    private static final String splitRegex = "\\r\n|\\n|\\.|\\。|\\！|\\!|\\?|\\？|\\、|\\;|\\；|\\'|\\\"|\\’|\\“|\\\\|\\<|\\>|\\《|\\》|\\{|\\}|\\|";

    /**
     * @Author  zhushouyong
     * @Description 将字符串按照固定长度分隔
     * @Date 2019-7-31 22:18
     * @Param [str, splitLen]
     * @return java.lang.String[]
    **/
    public static String[] splitStrByLength(String str, int splitLen) {
        if(null == str || splitLen <= 0) {
            return null;
        }
        if(str.length() < splitLen) {
            return new String[]{str};
        }
        int count = str.length() / splitLen + (str.length() % splitLen == 0 ? 0 : 1);
        String[] array = new String[count];
        for (int i = 0; i < count; i++) {
            if (str.length() <= splitLen) {
                array[i] = str;
            } else {
                array[i] = str.substring(0, splitLen);
                str = str.substring(splitLen);
            }
        }
        return array;
    }

    /**
     * @Author  zhushouyong
     * @Description 按照指定长度进行切割字符串，同时保证按照标点符号进行智能分隔，不会将语句断开
     * @Date 2019-7-31 22:14
     * @Param [str, splitLen]
     * @return java.lang.String[]
    **/
    public static String[] splitStr(String str,int splitLen){
        if(null == str || splitLen <= 0) {
            return null;
        }
        if(str.length() < splitLen) {
            return new String[]{str};
        }
        String[] array = str.split(splitRegex);//按照正则表达式规则进行分隔
        List<String> resultList = new ArrayList<String>();//存储返回结果
        StringBuilder stringBuilder = new StringBuilder();//存储每个元素
        for(int i=0,l=array.length;i<l;i++){
            if(array[i].length() <= splitLen){//当前字符串小于指定长度，则继续拼接
                if(stringBuilder.length() + array[i].length() > splitLen){//长度大于指定长度，则加入到结果中，同时新创建一个存储元素
                    resultList.add(stringBuilder.toString());
                    stringBuilder = new StringBuilder();
                }
                stringBuilder.append(array[i]);
            }else{//当前字符串大于指定长度，则先拼接，再分隔
                int subIndex = array[i].length()-stringBuilder.length();
                if(subIndex > splitLen){
                    subIndex = splitLen;
                }
                stringBuilder.append(array[i].substring(0,subIndex));
                resultList.add(stringBuilder.toString());
                stringBuilder = new StringBuilder();
                String[] subArray = splitStrByLength(array[i].substring(subIndex,array[i].length()),splitLen);//将剩下的部分进行按照指定长度进行分隔
                for(int j=0,s=subArray.length; j<s; j++){
                    if(subArray[j].length() >= splitLen){
                        resultList.add(subArray[j]);
                    }else {
                        stringBuilder.append(subArray[j]);
                    }
                }
            }
            if(i == (l-1)){
                resultList.add(stringBuilder.toString());
            }

        }
        String[] resultArray = new String[resultList.size()];
        resultList.toArray(resultArray);
        return resultArray;
    }

    /**
     * @Author  zhushouyong
     * @Description 过滤字符串，防止Json字符串冲突
     * @Date 2019-7-31 22:55
     * @Param [str]
     * @return java.lang.String
    **/
    public static String strFilter4JsonConflict(String str){
        if(str.indexOf("\"") > -1){
            str = str.replace("\"","\\\"");
        }
        return str;
    }


    public static String readStream(InputStream in) {
        ByteArrayOutputStream baos = null;
        String content = null;
        try {
            //<1>创建字节数组输出流，用来输出读取到的内容
            baos = new ByteArrayOutputStream();
            //<2>创建缓存大小
            byte[] buffer = new byte[1024]; // 1KB
            //每次读取到内容的长度
            int len = -1;
            //<3>开始读取输入流中的内容
            while ((len = in.read(buffer)) != -1) { //当等于-1说明没有数据可以读取了
                baos.write(buffer, 0, len);   //把读取到的内容写到输出流中
            }
            //<4> 把字节数组转换为字符串
            content = baos.toString("utf-8");


        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            try {
                //<5>关闭输入流和输出流
                in.close();
                baos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //<6>返回字符串结果
        return content;
    }

}