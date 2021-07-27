package com.ssm.test;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;

/**
 * @author zf
 * @date 2021/7/26 16:53
 */
public class HttpFIleUtil {

    /**
     * 通过附件url返回附件InputStream
     * @param path
     * @return
     */
    public static InputStream inputStream(String path) {
        URL url = null;
        InputStream is =null;
        try {
            url = new URL(path);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();//利用HttpURLConnection对象,我们可以从网络中获取网页数据.
            conn.setDoInput(true);
            conn.connect();
            is = conn.getInputStream();	//得到网络返回的输入流
        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return is;
    }
}
