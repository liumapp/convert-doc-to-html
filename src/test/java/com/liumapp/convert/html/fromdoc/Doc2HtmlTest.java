package com.liumapp.convert.html.fromdoc;

import org.junit.Test;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

/**
 * file Doc2HtmlTest.java
 * author liumapp
 * github https://github.com/liumapp
 * email liumapp.com@gmail.com
 * homepage http://www.liumapp.com
 * date 2018/9/27
 */
public class Doc2HtmlTest {

    private String dataPath = "/usr/local/tomcat/project/convert-doc-to-html/data/";

    @Test
    public void testConvert () throws IOException, TransformerException, ParserConfigurationException {
        Doc2Html doc2Html = new Doc2Html();
        doc2Html.setSavePath(dataPath + "/html/");
        String html = doc2Html.convert(new File(dataPath + "/doc/test.doc"));
        BufferedWriter bw = new BufferedWriter(new FileWriter(dataPath + "/html/test.html"));
        bw.write(html);
    }

}
