# convert-doc-to-html
Using xdocreport and aspose.word to convert doc to html . 

## how to use

add dependency by maven 

    <dependency>
        <groupId>com.liumapp.convert.html.fromdoc</groupId>
        <artifactId>convert-doc-to-html</artifactId>
        <version>v1.1.0</version>
    </dependency>

example :

    Doc2Html doc2Html = new Doc2Html();
    doc2Html.setSavePath(dataPath + "/html/");
    String html = doc2Html.convert(new File(dataPath + "/doc/test.doc"));
    BufferedWriter bw = new BufferedWriter(new FileWriter(dataPath + "/html/test.html"));
    bw.write(html);
    
    