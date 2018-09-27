package com.liumapp.convert.html.fromdoc;

import java.io.*;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import org.w3c.dom.Document;

/**
 * file Doc2Html.java
 * author liumapp
 * github https://github.com/liumapp
 * email liumapp.com@gmail.com
 * homepage http://www.liumapp.com
 * date 2018/9/27
 */
public class Doc2Html {

    private String savePath;

    public String getSavePath() {
        return savePath;
    }

    public Doc2Html setSavePath(String savePath) {
        this.savePath = savePath;
        return this;
    }

    public String convert (File file) throws ParserConfigurationException, IOException, TransformerException {
        InputStream input = new FileInputStream(file);
        String filename = file.getName();
        String suffix = filename.substring(filename.indexOf(".") + 1);
        if ("docx".equals(suffix)) {
            String content = parseDocxToHtml(file);
            return content;
        }
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
            HWPFDocument wordDocument = new HWPFDocument(input);
            wordToHtmlConverter.processDocument(wordDocument);
        }

        String content = conversion(wordToHtmlConverter);
        return content;
    }

    private String parseDocxToHtml(File file) throws IOException {
        if (!file.exists()) {
            return null;
        }
        if (file.getName().endsWith(".docx") || file.getName().endsWith(".DOCX")) {
            InputStream in = new FileInputStream(file);
            XWPFDocument document = new XWPFDocument(in);

            File imageFolderFile = new File(savePath);
            XHTMLOptions options = XHTMLOptions.create().URIResolver(
                    new FileURIResolver(imageFolderFile));
            options.setExtractor(new FileImageExtractor(imageFolderFile));
            options.setIgnoreStylesIfUnused(false);
            options.setFragment(true);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            XHTMLConverter.getInstance().convert(document, baos, options);
            String content = baos.toString();
            baos.close();
            return content;
        } else {
            System.out.println("Enter only MS Office 2007+ files");
        }
        return null;
    }

    private String conversion(WordToHtmlConverter wordToHtmlConverter) throws TransformerException, IOException {
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();
        String content = new String(outStream.toByteArray());
        return content;
    }

}
