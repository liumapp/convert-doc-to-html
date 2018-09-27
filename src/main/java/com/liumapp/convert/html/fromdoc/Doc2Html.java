package com.liumapp.convert.html.fromdoc;

import java.io.*;
import java.util.List;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * file Doc2Html.java
 * author liumapp
 * github https://github.com/liumapp
 * email liumapp.com@gmail.com
 * homepage http://www.liumapp.com
 * date 2018/9/27
 */
public class Doc2Html {

    public void convert() {
        String fileName = file.getOriginalFilename();
        SaveFile.savePic(file.getInputStream(), fileName);
        InputStream input = new FileInputStream(path + fileName);
        String suffix = fileName.substring(fileName.indexOf(".") + 1);// //截取文件格式名
        if ("docx".equals(suffix)) {
            String content = parseDocxToHtml(fileName);
            return content;
        }
        //实例化WordToHtmlConverter，为图片等资源文件做准备
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            public String savePicture(byte[] content, PictureType pictureType,
                                      String suggestedName, float widthInches, float heightInches) {
                return suggestedName;
            }
        });
        if ("doc".equals(suffix.toLowerCase())) {
            // doc
            HWPFDocument wordDocument = new HWPFDocument(input);
            wordToHtmlConverter.processDocument(wordDocument);
            //处理图片，会在同目录下生成并保存图片
            List pics = wordDocument.getPicturesTable().getAllPictures();
            if (pics != null) {
                for (int i = 0; i < pics.size(); i++) {
                    Picture pic = (Picture) pics.get(i);
                    try {
                        pic.writeImageContent(new FileOutputStream(path
                                + pic.suggestFullFileName()));
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        // 转换
        String content = conversion(wordToHtmlConverter);
        RespInfo respInfo = new RespInfo();
        if (content != null) {
            respInfo.setContent(content);
            respInfo.setMessage("success");
            respInfo.setStatus(InfoCode.SUCCESS);
        } else {
            respInfo.setMessage("error");
            respInfo.setStatus(InfoCode.ERROR);
        }
        return JSON.toJSONString(respInfo);
    }

}
