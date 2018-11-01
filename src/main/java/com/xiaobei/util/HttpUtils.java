package com.xiaobei.util;

import com.alibaba.fastjson.JSON;
import com.google.common.collect.Maps;
import org.apache.http.HttpResponse;
import org.apache.http.NameValuePair;
import org.apache.http.client.HttpClient;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.message.BasicNameValuePair;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Http工具
 *
 * @Author Lenovo
 * @Created 2016-06-05 19:29
 */
@SuppressWarnings("rawtypes")
public class HttpUtils {
    private final String USER_AGENT = "Mozilla/5.0";

    /**
     * http-请求  <br>
     *     注：TODO 未添加请求参数设置
     * @param url  接口地址
     * @param requestJsonParams 接口参数
     * @return  请求结果
     */
	public Map sendGet(String url, Map<String, String> requestJsonParams,  List<String> keys) throws Exception {

        HttpClient client = new DefaultHttpClient();
        HttpGet request = new HttpGet(url);

        // 添加请求头
        request.addHeader("User-Agent", USER_AGENT);
        HttpResponse response = client.execute(request);

        System.out.println("\n正在发送请求到接口地址：  " + url);
        System.out.println("requestJson值：" + requestJsonParams.get("requestJson"));
        System.out.println("responseStatus状态：" +   response.getStatusLine().getStatusCode());

        // 发起Post请求
        BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));

        // 读取响应报文
        StringBuffer result = new StringBuffer();
        String line = "";
        while ((line = rd.readLine()) != null) {
            result.append(line);
        }

        // 返回响应报文
        return JSON.parseObject(result.toString(), Map.class);

    }

    /**
     * http-post请求
     * @param url  接口地址
     * @param requestJsonParams 接口参数
     * @return  {url:"", request: "", response: ""}
     */
    public Map sendPost(String url, Map<String, String> requestJsonParams) throws Exception {

        HttpClient client = new DefaultHttpClient();
        HttpPost post = new HttpPost(url);

        // 添加请求头
        post.setHeader("User-Agent", USER_AGENT);
        // 添加请求参数
        List<NameValuePair> urlParameters = new ArrayList<NameValuePair>();
        // 添加请求参数
        for (Map.Entry<String, String> entry : requestJsonParams.entrySet()) {
        		String key = entry.getKey();
        		String value = entry.getValue();
        	   urlParameters.add(new BasicNameValuePair(key,  value));
               System.out.print("封装[" + key + "]值：" + value);
		}
         
        // 将请求参数设置到Entity中(此动作相当于将请求参数填充到Post请求的body中)
        post.setEntity(new UrlEncodedFormEntity(urlParameters, "UTF-8"));

        // 发起Post请求
        HttpResponse response = client.execute(post);
        System.out.println("\n正在发送请求到接口地址：  " + url);
        System.out.println("responseStatus状态：" +   response.getStatusLine().getStatusCode());

        // 读取响应报文
        BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
        StringBuffer result = new StringBuffer();
        String line = "";
        while ((line = rd.readLine()) != null) {
            result.append(line);
        }
        // 返回响应报文
        Map<String, String> resultMap = Maps.newHashMap();
        resultMap.put("url", url);
        resultMap.put("request", JSON.toJSONString(requestJsonParams) );
        resultMap.put("response", result.toString() );
        return resultMap;
    }

}
